from __future__ import annotations
import io
import re
import json
import tempfile
from pathlib import Path
from typing import Dict, Tuple, Optional, List

import pandas as pd
import pdfplumber
from openpyxl import load_workbook, Workbook

# 画像/OCR系（OCRが必要なPDF時にのみ使用）
import numpy as np
from pdf2image import convert_from_bytes
import cv2
import pytesseract
from pytesseract import Output


# ==============================
# 設定ロード
# ==============================
def load_config(path: str = "config.json") -> dict:
    p = Path(path)
    if not p.exists():
        # 最小デフォルト
        return {
            "excel_cell_map": {"sheet": "高圧", "法人名": "B1", "契約電力": "G12"},
            "excel_input": {"sheet_candidates": ["高圧", "Sheet1"]},
            "targets": {"法人名": {"regex": []}, "契約電力": {"regex": []}},
            "stop_keywords": ["会社", "株式会社", "有限会社", "医療法人", "病院", "クリニック"],
        }
    with p.open("r", encoding="utf-8") as f:
        cfg = json.load(f)

    # Tesseract/Poppler のパス指定があれば反映
    tess_cmd = cfg.get("tesseract_cmd")
    if tess_cmd:
        pytesseract.pytesseract.tesseract_cmd = tess_cmd

    return cfg


# ==============================
# 共通ユーティリティ
# ==============================
def _truncate_corporate_name(name: str, stop_words: List[str]) -> str:
    """法人名は『様』より左、かつ「会社/病院/クリニック」等の語尾で打ち切り。住所っぽい連番は除く。"""
    s = str(name).strip()
    s = s.split("様")[0]
    tails = ["株式会社", "有限会社", "医療法人", "学校法人", "合同会社", "合名会社", "合資会社", "病院", "クリニック", "会社"]
    for t in sorted(tails, key=len, reverse=True):
        if t in s:
            s = s[: s.find(t) + len(t)]
            break
    s = re.sub(r"[　 ]{2,}", " ", s)
    s = re.sub(r"[〒0-9０-９\-－ー–—−‐―/\.]{4,}.*$", "", s)
    return s.strip()


def _num_to_clean_str(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    s2 = s.replace(",", "")
    try:
        f = float(s2)
        if abs(f - round(f)) < 1e-9:
            return str(int(round(f)))
        return re.sub(r"0+$", "", f"{f}")
    except Exception:
        return s


# ==============================
# PDF: テキストPDFかを判定
# ==============================
def _is_text_pdf(pdf_bytes: bytes) -> Tuple[bool, str]:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if not pdf.pages:
                return False, ""
            page = pdf.pages[0]
            text = page.extract_text() or ""
            text = text.strip()
            return (len(text) > 20), text  # そこそこ文字があればテキストPDF扱い
    except Exception:
        return False, ""


# ==============================
# PDF: OCR（1〜3ページのみ）
# ==============================
def _ocr_text_and_words(pdf_bytes: bytes, cfg: dict, dpi: int = 300) -> Tuple[str, List[dict]]:
    pages_to_convert = 3  # 1〜3ページのみ
    poppler_path = cfg.get("poppler_path")  # 例: "C:/poppler-24.07.0/Library/bin"
    images = convert_from_bytes(
        pdf_bytes,
        dpi=dpi,
        first_page=1,
        last_page=pages_to_convert,
        poppler_path=poppler_path,
    )

    all_texts: List[str] = []
    all_words: List[dict] = []

    for img in images:
        # OpenCVに渡して前処理
        bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        # シャープ＋二値化（日本語の細線に強め）
        gray = cv2.bilateralFilter(gray, 9, 75, 75)
        thr = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                    cv2.THRESH_BINARY, 41, 9)

        data = pytesseract.image_to_data(thr, config="--psm 6 -l jpn+eng", output_type=Output.DICT)
        words = [dict(text=t, conf=c) for t, c in zip(data["text"], data["conf"]) if t and t.strip()]
        text_join = " ".join([w["text"] for w in words])
        all_texts.append(text_join)
        all_words.extend(words)

    return "\n".join(all_texts).strip(), all_words


# ==============================
# PDF: テキスト抽出→フィールド抽出
# ==============================
def _extract_fields_from_text(text: str, cfg: dict) -> Dict[str, str]:
    out: Dict[str, str] = {}

    # 法人名
    corp = None
    for pat in cfg.get("targets", {}).get("法人名", {}).get("regex", []):
        m = re.search(pat, text, flags=re.MULTILINE)
        if m:
            grp = m.group(1) if m.groups() else m.group(0)
            corp = grp.strip()
            break
    if corp:
        corp = _truncate_corporate_name(corp, cfg.get("stop_keywords", []))
        if corp:
            out["法人名"] = corp

    # 契約電力
    kw = None
    for pat in cfg.get("targets", {}).get("契約電力", {}).get("regex", []):
        m = re.search(pat, text, flags=re.MULTILINE)
        if m:
            kw = _num_to_clean_str(m.group(1))
            break
    if kw:
        out["契約電力"] = kw

    return out


# ==============================
# 外向け：PDF処理（自動判定）
# ==============================
def process_pdf_bytes(pdf_bytes: bytes, cfg: dict) -> Tuple[Dict[str, str], str, bool]:
    """戻り値: (抽出dict, 表示用テキスト, used_ocr)"""
    is_text, text0 = _is_text_pdf(pdf_bytes)
    if is_text:
        fields = _extract_fields_from_text(text0, cfg)
        return fields, text0, False

    # OCR
    text, _ = _ocr_text_and_words(pdf_bytes, cfg)
    fields = _extract_fields_from_text(text, cfg)
    return fields, text, True


# ==============================
# Excel: 抽出（ご指定どおり）
#  - B1 ← 「お客さま名」の右隣セル（会社/病院等で打ち切り）
#  - G12 ← 「契約電力(kW)」列の“最新月(YYYY年MM月)”の行の値
# ==============================
def _read_excel_to_df(xls_bytes: bytes, sheet_candidates: List[str]) -> pd.DataFrame:
    bio = io.BytesIO(xls_bytes)
    try:
        xl = pd.ExcelFile(bio)
        sheet = next((s for s in sheet_candidates if s in xl.sheet_names), xl.sheet_names[0])
        return xl.parse(sheet, header=None)
    except Exception:
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(xls_bytes), data_only=True)
        ws = wb[wb.sheetnames[0]]
        data = [[c for c in row] for row in ws.iter_rows(values_only=True)]
        return pd.DataFrame(data)


def process_excel_bytes(xls_bytes: bytes, cfg: dict) -> Tuple[Dict[str, str], pd.DataFrame]:
    excel_cfg = cfg.get("excel_input", {})
    sheet_candidates = excel_cfg.get("sheet_candidates", ["高圧", "Sheet1"])
    stop_words = cfg.get("stop_keywords", [])

    df = _read_excel_to_df(xls_bytes, sheet_candidates)
    H, W = df.shape
    result: Dict[str, str] = {}

    # 1) 「お客さま名」右隣
    corp_raw = None
    for r in range(min(H, 60)):
        for c in range(min(W, 50)):
            v = df.iat[r, c]
            if v is None:
                continue
            s = str(v)
            if ("お客さま名" in s) or ("お客様名" in s) or ("会社名" in s) or ("法人名" in s):
                cand = None
                if c + 1 < W and df.iat[r, c + 1] not in (None, ""):
                    cand = df.iat[r, c + 1]
                elif r + 1 < H and df.iat[r + 1, c] not in (None, ""):
                    cand = df.iat[r + 1, c]
                if cand not in (None, ""):
                    corp_raw = str(cand).strip()
                break
        if corp_raw:
            break
    if corp_raw:
        corp = _truncate_corporate_name(corp_raw, stop_words)
        if corp:
            result["法人名"] = corp

    # 2) 「契約電力(kW)」列の最新月
    power_col = None
    header_row = None
    for r in range(min(H, 30)):
        for c in range(W):
            val = df.iat[r, c]
            if val is None:
                continue
            s = str(val)
            if ("契約電力" in s) and (("kW" in s) or ("kw" in s) or ("ＫＷ" in s) or ("ｋＷ" in s) or ("kＷ" in s)):
                power_col, header_row = c, r
                break
        if power_col is not None:
            break

    # 月列を探す
    month_col = None
    if header_row is not None:
        for c in range(W):
            val = df.iat[header_row, c]
            if val is None:
                continue
            s = str(val)
            if ("月分" in s) or (s.strip() in ("月", "年月")):
                month_col = c
                break

    if month_col is None and header_row is not None:
        def count_month_like(ci: int) -> int:
            cnt = 0
            for r in range(header_row + 1, min(H, header_row + 40)):
                v = df.iat[r, ci]
                if v is None:
                    continue
                if re.search(r"20\d{2}年\s*\d{1,2}月", str(v)):
                    cnt += 1
            return cnt

        best_cnt, best_c = 0, None
        for c in range(min(W, 10)):
            cnt = count_month_like(c)
            if cnt > best_cnt:
                best_cnt, best_c = cnt, c
        if best_c is not None and best_cnt > 0:
            month_col = best_c

    def parse_month(s: str) -> Optional[Tuple[int, int]]:
        m = re.search(r"(20\d{2})\s*年\s*(\d{1,2})\s*月", s)
        if not m:
            return None
        return int(m.group(1)), int(m.group(2))

    latest_power = None
    if power_col is not None and header_row is not None:
        latest_key = (-1, -1)
        for r in range(header_row + 1, H):
            ym = None
            if month_col is not None:
                mv = df.iat[r, month_col]
                if mv is not None:
                    ym = parse_month(str(mv))
            if ym is None:
                continue

            pv = df.iat[r, power_col] if power_col < W else None
            if pv in (None, ""):
                continue

            if ym > latest_key:
                latest_key = ym
                latest_power = _num_to_clean_str(pv)

    if latest_power:
        result["契約電力"] = latest_power

    return result, df


# ==============================
# Excelテンプレ書き込み
# ==============================
def _ensure_template(path: Path, sheet: str):
    if path.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws["A1"] = "法人名"
    ws["A12"] = "契約電力"
    wb.save(path)


def write_to_excel(fields: Dict[str, str], cfg: dict, template_name: str = "template_output.xlsx") -> str:
    cell_map = cfg.get("excel_cell_map", {"sheet": "高圧", "法人名": "B1", "契約電力": "G12"})
    sheet = cell_map.get("sheet", "高圧")

    template_path = Path(template_name)
    _ensure_template(template_path, sheet)

    wb = load_workbook(template_path)
    if sheet not in wb.sheetnames:
        wb.create_sheet(title=sheet)
    ws = wb[sheet]

    # 書き込み
    if "法人名" in fields:
        ws[cell_map.get("法人名", "B1")] = fields["法人名"]
    if "契約電力" in fields:
        ws[cell_map.get("契約電力", "G12")] = fields["契約電力"]

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp_path = tmp.name
    tmp.close()
    return tmp_path

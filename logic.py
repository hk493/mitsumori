from __future__ import annotations
import io, re, json, tempfile
from pathlib import Path
from typing import Dict, Tuple, Optional, List

import numpy as np
import pandas as pd
import pdfplumber
import cv2
from pdf2image import convert_from_bytes
from openpyxl import load_workbook
from paddleocr import PaddleOCR


# ==============================
# 設定ロード
# ==============================
def load_config(path: str = "config.json") -> dict:
    p = Path(path)
    if not p.exists():
        # 最低限のデフォルト
        return {
            "poppler_path": None,
            "ocr": {"dpi": 400},
            "paddle": {"lang": "japan", "use_angle_cls": True},
            "excel_cell_map": {"sheet": "高圧", "法人名": "B1", "契約電力": "G12"},
            "stop_keywords": ["会社", "株式会社", "有限会社", "病院", "医院", "クリニック"],
            "targets": {"法人名": {"regex": []}, "契約電力": {"regex": []}},
            "excel_input": {"sheet_candidates": ["高圧", "Sheet1"]}
        }
    with p.open("r", encoding="utf-8") as f:
        return json.load(f)


# ==============================
# OCR 初期化
# ==============================
_OCR = None
def _get_ocr(lang: str = "japan", use_angle_cls: bool = True) -> PaddleOCR:
    global _OCR
    if _OCR is None:
        _OCR = PaddleOCR(lang=lang, use_angle_cls=use_angle_cls, show_log=False)
    return _OCR


# ==============================
# ユーティリティ
# ==============================
def _truncate_corporate_name(name: str, stop_words: List[str]) -> str:
    """法人名は『様』より左、かつ「会社/病院/クリニック」等で打ち切り。数字は除外。"""
    s = str(name).strip()
    # 数字が含まれる場合は住所等の可能性が高いので削る
    s = re.sub(r"[0-9０-９\-－ー–—‐―/\.]{3,}.*$", "", s)

    # 「様」より左を優先
    if "様" in s:
        s = s.split("様")[0]

    # 会社名の終端で打ち切り
    tails = ["株式会社", "有限会社", "医療法人", "学校法人", "合同会社", "合名会社", "合資会社", "病院", "医院", "クリニック", "会社"]
    for t in sorted(tails, key=len, reverse=True):
        if t in s:
            s = s[: s.find(t) + len(t)]
            break

    s = re.sub(r"[　 ]{2,}", " ", s)
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
# 前処理（標準）
# ==============================
def _deskew(gray):
    thr = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    coords = np.column_stack(np.where(thr == 0))
    if len(coords) < 100:
        return gray
    angle = cv2.minAreaRect(coords)[-1]
    angle = -(90 + angle) if angle < -45 else -angle
    h, w = gray.shape
    M = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    return cv2.warpAffine(gray, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)

def _preprocess_page(img):
    # RGB np.array -> 前処理（軽め）
    bgr = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 5, 50, 50)
    gray = cv2.convertScaleAbs(gray, alpha=1.3, beta=10)
    gray = _deskew(gray)
    bin_img = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                    cv2.THRESH_BINARY, 55, 7)
    return bin_img


# ==============================
# PDF：全ページOCR → テキスト連結
# ==============================
def _ocr_all_pages(pdf_bytes: bytes, cfg: dict) -> str:
    dpi = cfg.get("ocr", {}).get("dpi", 400)
    poppler = cfg.get("poppler_path")
    lang = cfg.get("paddle", {}).get("lang", "japan")
    use_angle_cls = cfg.get("paddle", {}).get("use_angle_cls", True)

    ocr = _get_ocr(lang=lang, use_angle_cls=use_angle_cls)
    pages = convert_from_bytes(pdf_bytes, dpi=dpi, poppler_path=poppler)
    texts: List[str] = []

    for idx, img in enumerate(pages, 1):
        arr = np.array(img)
        proc = _preprocess_page(arr)
        try:
            result = ocr.ocr(proc, cls=True)
        except Exception as e:
            print(f"OCR失敗({idx}): {e}")
            continue
        if not result or not isinstance(result, list) or not result[0]:
            continue

        page_texts: List[str] = []
        for line in result[0]:
            try:
                txt = line[1][0]
                if txt:
                    page_texts.append(txt)
            except Exception:
                continue
        texts.append("\n".join(page_texts))

    return "\n\n--- PAGE SPLIT ---\n\n".join(texts)


# ==============================
# PDF：抽出（法人名／契約電力）
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

    # 契約電力（kWhを誤検出しないように(?!h)）
    kw = None
    for pat in cfg.get("targets", {}).get("契約電力", {}).get("regex", []):
        m = re.search(pat, text, flags=re.MULTILINE | re.IGNORECASE)
        if m:
            # ラベル＋数値 or 数値のみ（第2グループ優先）
            if m.lastindex and m.lastindex >= 2 and m.group(2):
                kw = m.group(2)
            else:
                kw = m.group(1)
            kw = _num_to_clean_str(kw)
            break
    if kw:
        out["契約電力"] = kw

    return out


def process_pdf_bytes(pdf_bytes: bytes, cfg: dict) -> Tuple[Dict[str, str], str]:
    """常に全ページOCR → テキスト → 抽出dict を返す"""
    text = _ocr_all_pages(pdf_bytes, cfg) or ""
    fields = _extract_fields_from_text(text, cfg)
    return fields, text


# ==============================
# Excel：お客さま名の隣／最新月の契約電力
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


def process_excel_bytes(xls_bytes: bytes, cfg: dict):
    excel_cfg = cfg.get("excel_input", {})
    sheet_candidates = excel_cfg.get("sheet_candidates", ["高圧", "Sheet1"])
    stop_words = cfg.get("stop_keywords", [])

    df = _read_excel_to_df(xls_bytes, sheet_candidates)
    H, W = df.shape
    result: Dict[str, str] = {}

    # 1) 「お客さま名」右隣（なければ下のセル）
    corp_raw = None
    for r in range(min(H, 80)):
        for c in range(min(W, 80)):
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

    # 2) 契約電力(kW)列 & 月列
    power_col = None
    header_row = None
    for r in range(min(H, 50)):
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

    # 月列の推測
    month_col = None
    if header_row is not None:
        # 明示的ヘッダ
        for c in range(W):
            val = df.iat[header_row, c]
            if val is None:
                continue
            s = str(val)
            if ("月分" in s) or (s.strip() in ("月", "年月")):
                month_col = c
                break

        # パターンから最多一致の列を選ぶ
        if month_col is None:
            def count_month_like(ci: int) -> int:
                cnt = 0
                for r in range(header_row + 1, min(H, header_row + 60)):
                    v = df.iat[r, ci]
                    if v is None:
                        continue
                    if re.search(r"20\d{2}年\s*\d{1,2}月", str(v)):
                        cnt += 1
                return cnt

            best_cnt, best_c = 0, None
            for c in range(min(W, 15)):
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
    if power_col is not None and header_row is not None and month_col is not None:
        latest_key = (-1, -1)
        for r in range(header_row + 1, H):
            mv = df.iat[r, month_col]
            if mv is None:
                continue
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
# 既存テンプレへ書き込み（必須）
# ==============================
def write_to_excel(fields: Dict[str, str], cfg: dict, template_name: str = "template_output.xlsx") -> str:
    cell_map = cfg.get("excel_cell_map", {"sheet": "高圧", "法人名": "B1", "契約電力": "G12"})
    sheet = cell_map.get("sheet", "高圧")

    template_path = Path(template_name)
    if not template_path.exists():
        # 既存テンプレが必須
        raise FileNotFoundError("template_output.xlsx が見つかりません。プロジェクト直下に置いてから再実行してください。")

    wb = load_workbook(template_path)
    if sheet not in wb.sheetnames:
        wb.create_sheet(title=sheet)
    ws = wb[sheet]

    if "法人名" in fields:
        ws[cell_map.get("法人名", "B1")] = fields["法人名"]
    if "契約電力" in fields:
        ws[cell_map.get("契約電力", "G12")] = fields["契約電力"]

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    tmp_path = tmp.name
    return tmp_path

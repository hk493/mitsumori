from __future__ import annotations
import re, json, tempfile
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple

import pdfplumber
from openpyxl import load_workbook


# =========================
# 設定
# =========================
def load_config(path: str | Path) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# =========================
# 共通ユーティリティ
# =========================
def _apply_postprocess(value: Optional[str], spec: Dict[str, Any]) -> Any:
    """後処理は最小限。数値化せず、読み取った文字列をそのまま返す。"""
    if value is None:
        return None
    v = value
    if spec.get("strip"):
        v = v.strip(spec.get("trim_chars") or None)
    return v


def _try_regex(text: str, patterns: List[str]) -> Optional[str]:
    for pat in patterns:
        m = re.search(pat, text, flags=re.IGNORECASE)
        if m:
            return (m.group(1) or "").strip()
    return None


def _truncate_corporate_name(raw: str, end_markers: List[str]) -> str:
    """
    法人名の末尾を終端キーワード（病院/医院/クリニック/会社など）で切る。
    もっとも左（早い位置）でヒットした marker の直後で切り、余分な空白を除去。
    """
    if not raw:
        return raw
    pos = None
    for mk in end_markers:
        i = raw.find(mk)
        if i != -1:
            end = i + len(mk)
            pos = end if (pos is None or end < pos) else pos
    return raw[:pos].rstrip(" 　\t") if pos else raw


# =========================
# PDFテキスト & 単語抽出
# =========================
def _extract_text_and_words(pdf_bytes: bytes) -> Tuple[str, List[Dict[str, Any]]]:
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        texts = []
        words: List[Dict[str, Any]] = []
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ""
            texts.append(t)
            if i == 0:
                words = p.extract_words(use_text_flow=True, keep_blank_chars=False)
        return "\n".join(texts), words


# =========================
# 法人名：regex → 幾何（「様」直上/左近傍）→ 終端キーワードで切り
# =========================
def _extract_corporate(text: str, words: List[Dict[str, Any]], spec: Dict[str, Any]) -> Optional[str]:
    corp = _try_regex(text, spec.get("regex_patterns", []))

    if not corp:
        # 幾何ヒントで拾う
        hints = spec.get("geom_hints", {}) or {}
        anchors = hints.get("anchor_texts", ["様"])
        x_tol = int(hints.get("x_tolerance", 40))
        y_up = int(hints.get("y_up_range", 60))

        if words:
            anchor_words = [w for w in words if any(a in w.get("text", "") for a in anchors)]
            if anchor_words:
                anchor = sorted(anchor_words, key=lambda w: (w.get("top", 1e9), w.get("x0", 1e9)))[0]
                ax0, atop = anchor.get("x0", 0), anchor.get("top", 0)

                candidates = [
                    w for w in words
                    if atop - y_up <= w.get("top", 0) < atop and abs(w.get("x0", 0) - ax0) <= x_tol
                ]
                if not candidates:
                    same_line_left = [w for w in words if abs(w.get("top", 0) - atop) <= 5 and w.get("x1", 0) <= ax0]
                    same_line_left = sorted(same_line_left, key=lambda w: w.get("x0", 0))
                    corp = "".join([w.get("text", "") for w in same_line_left]).strip() or None
                else:
                    candidates = sorted(candidates, key=lambda w: (w.get("top", 0), w.get("x0", 0)))
                    corp = "".join([w.get("text", "") for w in candidates]).strip() or None

    if corp:
        markers = spec.get("corp_end_markers", [])
        corp = _truncate_corporate_name(corp, markers)
        corp = _apply_postprocess(corp, spec.get("postprocess", {}))
    return corp


# =========================
# 契約電力：主契約の右隣（同一行テキスト）を最優先
# =========================
def _extract_contract_power_by_wordline(
    pdf_bytes: bytes,
    keywords: List[str] = ["主契約", "主 契約"],
    line_tol: float = 3.5
) -> Optional[str]:
    def norm(s: str) -> str:
        return (s or "").replace(" ", "").replace("　", "")

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        if not pdf.pages:
            return None
        p = pdf.pages[0]
        words = p.extract_words(use_text_flow=True, keep_blank_chars=False) or []
        if not words:
            return None

        words = sorted(words, key=lambda w: (w.get("top", 0.0), w.get("x0", 0.0)))
        lines, current, last_top = [], [], None
        for w in words:
            top = float(w.get("top", 0.0))
            if last_top is None or abs(top - last_top) <= line_tol:
                current.append(w)
                last_top = top if last_top is None else (last_top + top) / 2.0
            else:
                lines.append(current)
                current, last_top = [w], top
        if current:
            lines.append(current)

        for line in lines:
            if not any(any(k in norm(w.get("text", "")) for k in keywords) for w in line):
                continue
            line = sorted(line, key=lambda ww: ww.get("x0", 0.0))
            concat_text = "".join(ww.get("text", "") for ww in line)
            m = re.search(r"([0-9,]+(?:\.\d+)?)\s*(?:kW|ＫＷ|kw)", concat_text, flags=re.IGNORECASE)
            if m:
                return m.group(1)
    return None


# =========================
# 基本料金 × kW／kWh 列（ヘッダ列のX範囲で同列の値を取得）
# =========================
def _extract_kw_from_basic_row_by_words(
    pdf_bytes: bytes,
    basic_row_keywords: List[str],
    kw_header_keywords: List[str],
    line_tol: float = 3.5,
    col_margin: float = 8.0
) -> Optional[str]:
    def norm(s: str) -> str:
        return (s or "").replace(" ", "").replace("　", "")

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        if not pdf.pages:
            return None
        p = pdf.pages[0]
        words = p.extract_words(use_text_flow=True, keep_blank_chars=False) or []
        if not words:
            return None

        # 行グループ化
        words = sorted(words, key=lambda w: (w.get("top", 0.0), w.get("x0", 0.0)))
        lines, current, last_top = [], [], None
        for w in words:
            top = float(w.get("top", 0.0))
            if last_top is None or abs(top - last_top) <= line_tol:
                current.append(w)
                last_top = top if last_top is None else (last_top + top) / 2.0
            else:
                lines.append(current)
                current, last_top = [w], top
        if current:
            lines.append(current)

        # 1) ヘッダ行から kW 列の x 範囲を推定
        kw_x0, kw_x1 = None, None
        for line in lines:
            for w in line:
                t = norm(w.get("text", ""))
                if any(norm(h) in t for h in kw_header_keywords):
                    xs = sorted([ww.get("x0", 0.0) for ww in line] + [ww.get("x1", 0.0) for ww in line])
                    x_center = (w.get("x0", 0.0) + w.get("x1", 0.0)) / 2
                    lefts = [x for x in xs if x <= x_center]
                    rights = [x for x in xs if x >= x_center]
                    kw_x0 = (lefts[-2] if len(lefts) >= 2 else w.get("x0", 0.0)) - col_margin
                    kw_x1 = (rights[1] if len(rights) >= 2 else w.get("x1", 0.0)) + col_margin
                    break
            if kw_x0 is not None:
                break
        if kw_x0 is None:
            return None

        # 2) 基本料金行で、その列範囲に入る数値を取得
        for line in lines:
            if not any(norm(k) in norm(w.get("text", "")) for k in basic_row_keywords for w in line):
                continue
            for w in sorted(line, key=lambda ww: ww.get("x0", 0.0)):
                cx = (w.get("x0", 0.0) + w.get("x1", 0.0)) / 2
                if kw_x0 <= cx <= kw_x1:
                    m = re.search(r"([0-9,]+(?:\.\d+)?)", w.get("text", ""))
                    if m:
                        return m.group(1)
        return None


# =========================
# 従来テーブル抽出（extract_tables 経由）
# =========================
def _extract_contract_power_from_tables(pdf_bytes: bytes, table_spec: Dict[str, Any]) -> Optional[str]:
    row_labels = [s.replace(" ", "").replace("　", "") for s in table_spec.get("row_label_keywords", [])]
    neighbor_offset = table_spec.get("neighbor_col_offset", None)
    kw_headers = table_spec.get("header_kw_for_kw_col", [])
    fallback_offset = int(table_spec.get("fallback_col_offset_from_label", 2))

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        if not pdf.pages:
            return None
        p = pdf.pages[0]
        tables = p.extract_tables() or []

        for tbl in tables:
            norm = [[(c or "").strip() for c in row] for row in tbl if any(row)]
            if not norm:
                continue
            header = norm[0]

            def is_kw_header(s: str) -> bool:
                s2 = s.replace(" ", "").replace("　", "")
                return any(k.replace(" ", "") in s2 for k in kw_headers)

            kw_col_idx = next((j for j, h in enumerate(header) if is_kw_header(h)), None)

            for row in norm[1:]:
                if not row:
                    continue
                label = (row[0] or "").replace(" ", "").replace("　", "")
                if not any(lbl in label for lbl in row_labels):
                    continue

                if neighbor_offset is not None:
                    target_idx = 0 + int(neighbor_offset)
                    if target_idx < len(row) and row[target_idx]:
                        return row[target_idx]

                if kw_col_idx is not None and kw_col_idx < len(row) and row[kw_col_idx]:
                    return row[kw_col_idx]

                target_idx = 0 + fallback_offset
                if target_idx < len(row) and row[target_idx]:
                    return row[target_idx]
    return None


# =========================
# 契約電力：抽出順序の総合関数（すべて文字列で返す）
# =========================
def _extract_contract_power(pdf_bytes: bytes, text: str, spec: Dict[str, Any]) -> Optional[str]:
    v_line = _extract_contract_power_by_wordline(pdf_bytes)
    if v_line:
        return v_line

    table = spec.get("table", {}) or {}
    v_basic = _extract_kw_from_basic_row_by_words(
        pdf_bytes,
        basic_row_keywords=table.get("basic_row_keywords", []),
        kw_header_keywords=table.get("kw_header_keywords", []),
    )
    if v_basic:
        return v_basic

    if table:
        v_tbl = _extract_contract_power_from_tables(pdf_bytes, table)
        if v_tbl:
            return v_tbl

    v_re = _try_regex(text, spec.get("regex_patterns", []))
    return v_re


# =========================
# 公開：1PDF → 抽出辞書
# =========================
def extract_data(pdf_bytes: bytes, config: Dict[str, Any]) -> Dict[str, Any]:
    text, words = _extract_text_and_words(pdf_bytes)
    out: Dict[str, Any] = {"raw_text": text}

    corp_spec = config["targets"]["法人名"]
    corp = _extract_corporate(text, words, corp_spec)
    out["法人名"] = corp

    pow_spec = config["targets"]["契約電力"]
    power = _extract_contract_power(pdf_bytes, text, pow_spec)
    out["契約電力"] = _apply_postprocess(power, pow_spec.get("postprocess", {}))

    return out


# =========================
# 既存テンプレに“だけ”書き込む
# =========================
def write_to_excel_with_mapping(data: Dict[str, Any], config: Dict[str, Any],
                                template_path: str | Path) -> str:
    wb = load_workbook(template_path)
    for key, spec in config.get("targets", {}).items():
        sheet = spec.get("sheet", "高圧")
        cell = spec.get("cell", "A1")
        if sheet in wb.sheetnames and (key in data) and (data[key] not in (None, "")):
            wb[sheet][cell] = data[key]

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(tmp.name)
    path = tmp.name
    tmp.close()
    return path


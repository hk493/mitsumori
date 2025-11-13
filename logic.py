import io
import re
import json
import unicodedata
from copy import deepcopy
from typing import Dict, Any, List, Optional, Tuple

import pdfplumber
from openpyxl import load_workbook

# OCR は「OCRで解析」ボタン時のみ（入っていなければスキップ）
try:
    from pdf2image import convert_from_bytes
    import pytesseract
    from pytesseract import Output as TesseractOutput
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False


# ===================== 設定ロード =====================
def load_config(path_or_file) -> Dict[str, Any]:
    """config.json を dict で返す（ファイルパス or ファイルオブジェクト）"""
    if hasattr(path_or_file, "read"):
        return json.load(path_or_file)
    with open(path_or_file, "r", encoding="utf-8") as f:
        return json.load(f)


# ===================== 低レベル抽出（PDF→テキスト/単語） =====================
def _extract_text_with_pdfplumber(pdf_bytes: bytes) -> Tuple[str, List[Dict[str, Any]]]:
    texts: List[str] = []
    words_all: List[Dict[str, Any]] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or "")
            try:
                words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
            except Exception:
                words = []
            for w in words:
                w["page"] = page.page_number
            words_all.extend(words)
    return "\n".join(texts), words_all


def _extract_text_with_ocr(pdf_bytes: bytes, ocr_cfg: Dict[str, Any]) -> Tuple[str, List[Dict[str, Any]]]:
    if not OCR_AVAILABLE:
        raise RuntimeError("OCRモジュール（pdf2image/pytesseract）が未導入です。")

    lang = ocr_cfg.get("lang", "jpn+eng")
    psm  = ocr_cfg.get("psm", 6)
    oem  = ocr_cfg.get("oem", 3)
    dpi  = ocr_cfg.get("dpi", 300)
    tcmd = ocr_cfg.get("tesseract_cmd")
    ppop = ocr_cfg.get("poppler_path")
    if tcmd:
        pytesseract.pytesseract.tesseract_cmd = tcmd

    images = convert_from_bytes(pdf_bytes, dpi=dpi, poppler_path=ppop)
    all_text: List[str] = []
    words_all: List[Dict[str, Any]] = []

    for page_idx, img in enumerate(images, start=1):
        data = pytesseract.image_to_data(
            img, lang=lang, config=f"--psm {psm} --oem {oem}", output_type=TesseractOutput.DICT
        )
        tokens: List[str] = []
        for i in range(len(data["text"])):
            t = (data["text"][i] or "").strip()
            if not t:
                continue
            tokens.append(t)
            words_all.append({
                "text": t,
                "x0": data["left"][i],
                "top": data["top"][i],
                "x1": data["left"][i] + data["width"][i],
                "bottom": data["top"][i] + data["height"][i],
                "page": page_idx,
            })
        all_text.append(" ".join(tokens))

    return "\n".join(all_text), words_all


# ===================== 正規化ユーティリティ =====================
def _normalize_token(s: str, ncfg: dict) -> str:
    if not isinstance(s, str):
        return s
    # 全角→半角（幅・記号ゆらぎ吸収）
    if ncfg.get("zen_to_han", True):
        s = unicodedata.normalize("NFKC", s)
    # 単位・記号の別表記を揃える
    for k, v in (ncfg.get("unit_aliases") or {}).items():
        s = s.replace(k, v)
    for k, v in (ncfg.get("char_map") or {}).items():
        s = s.replace(k, v)
    # 余分な空白
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _normalize_text_and_words(raw_text: str, words: list, config: dict):
    ncfg = config.get("normalize", {})
    norm_text = _normalize_token(raw_text or "", ncfg)
    norm_words = []
    for w in words or []:
        w = dict(w)
        t = w.get("text")
        if isinstance(t, str):
            w["text"] = _normalize_token(t, ncfg)
        norm_words.append(w)
    return norm_text, norm_words


# ===================== 汎用ユーティリティ =====================
def _post(v: Optional[str], pp: Dict[str, Any]):
    """トリム・数値化・レンジチェックなどの後処理"""
    if v is None:
        return None
    if pp.get("strip"):
        v = v.strip(pp.get("trim_chars")) if pp.get("trim_chars") else v.strip()
    if pp.get("to_number"):
        try:
            v = float(str(v).replace(",", ""))
        except Exception:
            return None
        # 異常値ガード
        rng = pp.get("valid_range")
        if isinstance(rng, list) and len(rng) == 2:
            lo, hi = rng
            try:
                if v < lo or v > hi:
                    return None
            except Exception:
                pass
    return v


def _regex_first(text: str, patterns: List[str]) -> Optional[str]:
    for p in patterns:
        m = re.search(p, text, flags=re.MULTILINE | re.IGNORECASE)
        if m:
            g = m.group(1) if m.lastindex else m.group(0)
            return (g or "").strip()
    return None


def _find_near_anchor(words: List[Dict[str, Any]], anchor_texts: List[str], x_tol: float, y_up: float) -> Optional[str]:
    """アンカー語（例：様）の左・上近傍から連結したテキストを推定"""
    anchors = [w for w in words if w.get("text") in anchor_texts]
    if not anchors:
        return None
    a = anchors[0]
    ax0, atop = a["x0"], a["top"]
    cands: List[str] = []
    for w in words:
        if w.get("page") != a.get("page"):
            continue
        # アンカー左上矩形（誤差吸収）
        if (0 <= (ax0 - w["x1"]) <= x_tol) and (0 < (atop - w["bottom"]) <= y_up):
            cands.append(w["text"])
        # 同一行で左側
        same_line = abs(w["top"] - a["top"]) <= max(3, 0.02 * a["top"]) and w["x1"] <= a["x0"]
        if same_line:
            cands.append(w["text"])
    if not cands:
        return None
    return "".join(cands[-5:]).strip()


# ===================== 帳票分類 =====================
def _deep_merge(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    r = deepcopy(a)
    for k, v in b.items():
        if isinstance(v, dict) and isinstance(r.get(k), dict):
            r[k] = _deep_merge(r[k], v)
        else:
            r[k] = v
    return r


def classify_form(raw_text: str, config: Dict[str, Any]):
    default_targets = config.get("default_targets", {})
    for f in config.get("forms", []):
        rule = f.get("detect", {})
        all_of = rule.get("all_of", [])
        any_of = rule.get("any_of", [])
        none_of = rule.get("none_of", [])
        if any(x not in raw_text for x in all_of):
            continue
        if any_of and not any(x in raw_text for x in any_of):
            continue
        if any(x in raw_text for x in none_of):
            continue
        return f.get("id", "default"), _deep_merge(default_targets, f.get("targets", {}))
    return "default", default_targets


# ===================== kW/kWh 判定と数値抽出 =====================
def _is_kw_only(s: str) -> bool:
    """kW を含むが kWh は含まない（正規化済み前提）"""
    return bool(re.search(r"kW(?!h)", s, flags=re.I))


def _is_kwh(s: str) -> bool:
    return "kwh" in s.lower()


def _first_number(s: str) -> Optional[str]:
    m = re.search(r"([0-9]+(?:,[0-9]{3})*(?:\.[0-9]+)?)", s)
    return m.group(1) if m else None


# ===================== 値の抽出（法人名／契約電力） =====================
def _extract_corp(raw_text: str, words: List[Dict[str, Any]], tcfg: Dict[str, Any]) -> Optional[str]:
    # 1) 正規表現
    v = _regex_first(raw_text, tcfg.get("regex_patterns", []))
    if v:
        return _post(v, tcfg.get("postprocess", {}))
    # 2) アンカー近傍（様/御中 など）
    gh = tcfg.get("geom_hints", {})
    if gh:
        v = _find_near_anchor(
            words,
            gh.get("anchor_texts", ["様"]),
            gh.get("x_tolerance", 40),
            gh.get("y_up_range", 60),
        )
        if v:
            return _post(v, tcfg.get("postprocess", {}))
    return None


def _kw_from_tables(pdf_bytes: bytes, tcfg: Dict[str, Any], config: Dict[str, Any]) -> Optional[str]:
    """表から契約電力（kWのみ）を抽出。kWh を明確に除外。"""
    if not tcfg.get("table"):
        return None

    tb = tcfg["table"]
    ncfg = config.get("normalize", {})

    req = [x.lower() for x in (tb.get("unit_policy", {}).get("require_any") or [])]
    exc = [x.lower() for x in (tb.get("unit_policy", {}).get("exclude_any") or [])]

    headers = [h.lower() for h in (tb.get("header_kw_for_kw_col") or [])]
    rows_kw = tb.get("row_label_keywords", [])
    rows_rx = [re.compile(rx) for rx in (tb.get("row_label_regex") or [])]
    off = tb.get("fallback_col_offset_from_label", 2)

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                tables = page.extract_tables()
            except Exception:
                tables = []
            for tbl in tables or []:
                norm = [[_normalize_token((c or ""), ncfg) for c in row] for row in (tbl or []) if row]
                if not norm:
                    continue

                # ヘッダから kW 列（kWh混在を排除）
                kw_col = None
                if norm:
                    for j, cell in enumerate(norm[0]):
                        c = cell.lower()
                        if any(h in c for h in headers) and _is_kw_only(c):
                            kw_col = j
                            break

                # ラベル行（基本料金/契約電力/最大需要電力など）から値を拾う
                for row in norm:
                    label = row[0] if row else ""
                    label_ok = (any(k in label for k in rows_kw) or any(r.search(label) for r in rows_rx))
                    if not label_ok:
                        continue

                    # 1) 推定した kW 列を優先
                    if kw_col is not None and kw_col < len(row):
                        cell = row[kw_col]
                        cond_req = (("kw" in req) and _is_kw_only(cell)) or (not req)
                        cond_exc = (("kwh" in exc) and (not _is_kwh(cell))) or (not exc)
                        if cond_req and cond_exc and _is_kw_only(cell):
                            v = _first_number(cell)
                            if v:
                                return v

                    # 2) ラベルからの相対（横の横）
                    if len(row) > off:
                        cell = row[off]
                        cond_req = (("kw" in req) and _is_kw_only(cell)) or (not req)
                        cond_exc = (("kwh" in exc) and (not _is_kwh(cell))) or (not exc)
                        if cond_req and cond_exc and _is_kw_only(cell):
                            v = _first_number(cell)
                            if v:
                                return v

                # 3) テーブル全体スキャン（kWhを含むセルは除外）
                for row in norm:
                    for cell in row:
                        if _is_kw_only(cell) and not _is_kwh(cell):
                            v = _first_number(cell)
                            if v:
                                return v

    return None


def _kw_from_text(raw_text: str, tcfg: Dict[str, Any]) -> Optional[str]:
    """テキスト行から契約電力（kWのみ）を抽出。kWh を除外。"""
    for ln in [x.strip() for x in raw_text.splitlines() if x.strip()]:
        if _is_kw_only(ln) and not _is_kwh(ln):
            v = _first_number(ln)
            if v:
                return v
    # 最後に正規表現でフォールバック（kW(?!h) を使う）
    return _regex_first(raw_text, tcfg.get("regex_patterns", []))


def _extract_pair(raw_text: str, words: List[Dict[str, Any]], targets: Dict[str, Any],
                  pdf_bytes_for_tables: Optional[bytes], config: Dict[str, Any]):
    out: Dict[str, Any] = {}

    # 法人名
    corp_cfg = targets.get("法人名", {})
    out["法人名"] = _extract_corp(raw_text, words, corp_cfg)

    # 契約電力
    kw_cfg = targets.get("契約電力", {})
    kw_val = None
    if pdf_bytes_for_tables is not None:
        kw_val = _kw_from_tables(pdf_bytes_for_tables, kw_cfg, config)
    if not kw_val:
        kw_val = _kw_from_text(raw_text, kw_cfg)
    out["契約電力"] = _post(kw_val, kw_cfg.get("postprocess", {}))

    # （必要なら将来ここで他のターゲットも追加）
    return out


# ===================== 公開API（UIから呼ぶ） =====================
def extract_data_text(pdf_bytes: bytes, config: Dict[str, Any]) -> Dict[str, Any]:
    # テキスト抽出 → 正規化 → 帳票分類 → 値抽出
    raw_text, words = _extract_text_with_pdfplumber(pdf_bytes)
    raw_text, words = _normalize_text_and_words(raw_text, words, config)
    form_id, targets = classify_form(raw_text, config)
    data = _extract_pair(raw_text, words, targets, pdf_bytes, config)
    data.update({"raw_text": raw_text, "form_id": form_id})
    return data


def extract_data_ocr(pdf_bytes: bytes, config: Dict[str, Any]) -> Dict[str, Any]:
    # OCR抽出 → 正規化 → 帳票分類 → 値抽出（表解析はしない：テキストベースに統一）
    ocr_cfg = config.get("ocr", {"lang": "jpn+eng", "psm": 6, "oem": 3, "dpi": 300})
    raw_text, words = _extract_text_with_ocr(pdf_bytes, ocr_cfg)
    raw_text, words = _normalize_text_and_words(raw_text, words, config)
    form_id, targets = classify_form(raw_text, config)
    data = _extract_pair(raw_text, words, targets, None, config)
    data.update({"raw_text": raw_text, "form_id": form_id})
    return data


# ===================== Excel 書き込み（体裁保持・指定セルのみ） =====================
def _effective_targets(config: Dict[str, Any], form_id: Optional[str]) -> Dict[str, Any]:
    base = config.get("default_targets", {})
    if not form_id or form_id == "default":
        return base
    for f in config.get("forms", []):
        if f.get("id") == form_id:
            return _deep_merge(base, f.get("targets", {}))
    return base


def write_to_excel_with_mapping(data: Dict[str, Any], config: Dict[str, Any], template_path,
                                form_id: Optional[str] = None) -> bytes:
    """テンプレ（体裁保持）に指定セルだけ上書きして Bytes を返す"""
    if hasattr(template_path, "read"):
        wb = load_workbook(filename=template_path)
    else:
        wb = load_workbook(filename=str(template_path))

    targets = _effective_targets(config, form_id)
    for key, tcfg in targets.items():
        sheet = tcfg.get("sheet")
        cell  = tcfg.get("cell")
        if sheet and cell and (key in data) and (data[key] is not None):
            try:
                ws = wb[sheet]
                ws[cell] = data[key]
            except KeyError:
                # 指定シートが無い場合はスキップ
                continue

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

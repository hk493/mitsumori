import io, re, json
from copy import deepcopy
from typing import Dict, Any, List, Optional, Tuple

import pdfplumber
from openpyxl import load_workbook

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    from pytesseract import Output as TesseractOutput
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False

def load_config(path_or_file) -> Dict[str, Any]:
    if hasattr(path_or_file, \"read\"):
        return json.load(path_or_file)
    with open(path_or_file, \"r\", encoding=\"utf-8\") as f:
        return json.load(f)

def _extract_text_with_pdfplumber(pdf_bytes: bytes) -> Tuple[str, List[Dict[str, Any]]]:
    texts: List[str] = []
    words_all: List[Dict[str, Any]] = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            texts.append(page.extract_text() or \"\")
            try:
                words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
            except Exception:
                words = []
            for w in words:
                w[\"page\"] = page.page_number
            words_all.extend(words)
    return \"\n\".join(texts), words_all

def _extract_text_with_ocr(pdf_bytes: bytes, ocr_cfg: Dict[str, Any]) -> Tuple[str, List[Dict[str, Any]]]:
    if not OCR_AVAILABLE:
        raise RuntimeError(\"OCRモジュール（pdf2image/pytesseract）が未導入です。\")
    lang = ocr_cfg.get(\"lang\", \"jpn+eng\")
    psm  = ocr_cfg.get(\"psm\", 6)
    oem  = ocr_cfg.get(\"oem\", 3)
    dpi  = ocr_cfg.get(\"dpi\", 300)
    tcmd = ocr_cfg.get(\"tesseract_cmd\")
    ppop = ocr_cfg.get(\"poppler_path\")
    if tcmd:
        pytesseract.pytesseract.tesseract_cmd = tcmd
    images = convert_from_bytes(pdf_bytes, dpi=dpi, poppler_path=ppop)
    all_text: List[str] = []
    words_all: List[Dict[str, Any]] = []
    for page_idx, img in enumerate(images, start=1):
        data = pytesseract.image_to_data(
            img, lang=lang, config=f\"--psm {psm} --oem {oem}\", output_type=TesseractOutput.DICT
        )
        tokens: List[str] = []
        for i in range(len(data[\"text\"])):
            t = (data[\"text\"][i] or \"\").strip()
            if not t:
                continue
            tokens.append(t)
            words_all.append({
                \"text\": t,
                \"x0\": data[\"left\"][i],
                \"top\": data[\"top\"][i],
                \"x1\": data[\"left\"][i] + data[\"width\"][i],
                \"bottom\": data[\"top\"][i] + data[\"height\"][i],
                \"page\": page_idx,
            })
        all_text.append(\" \".join(tokens))
    return \"\n\".join(all_text), words_all

def _post(v: Optional[str], pp: Dict[str, Any]):
    if v is None:
        return None
    if pp.get(\"strip\"):
        v = v.strip(pp.get(\"trim_chars\")) if pp.get(\"trim_chars\") else v.strip()
    if pp.get(\"to_number\"):
        try:
            v = float(str(v).replace(\",\", \"\"))
        except Exception:
            v = None
    return v

def _regex_first(text: str, patterns: List[str]) -> Optional[str]:
    for p in patterns:
        m = re.search(p, text, flags=re.MULTILINE)
        if m:
            return (m.group(1) or m.group(0)).strip()
    return None

def _find_near_anchor(words: List[Dict[str, Any]], anchor_texts: List[str], x_tol: float, y_up: float) -> Optional[str]:
    anchors = [w for w in words if w.get(\"text\") in anchor_texts]
    if not anchors:
        return None
    a = anchors[0]
    ax0, atop = a[\"x0\"], a[\"top\"]
    cands: List[str] = []
    for w in words:
        if w.get(\"page\") != a.get(\"page\"):
            continue
        if (0 <= (ax0 - w[\"x1\"]) <= x_tol) and (0 < (atop - w[\"bottom\"]) <= y_up):
            cands.append(w[\"text\"])
        same_line = abs(w[\"top\"] - a[\"top\"]) <= max(3, 0.02 * a[\"top\"]) and w[\"x1\"] <= a[\"x0\"]
        if same_line:
            cands.append(w[\"text\"])
    if not cands:
        return None
    return \"\".join(cands[-5:]).strip()

def _deep_merge(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    r = deepcopy(a)
    for k, v in b.items():
        if isinstance(v, dict) and isinstance(r.get(k), dict):
            r[k] = _deep_merge(r[k], v)
        else:
            r[k] = v
    return r

def classify_form(raw_text: str, config: Dict[str, Any]):
    default_targets = config.get(\"default_targets\", {})
    for f in config.get(\"forms\", []):
        rule = f.get(\"detect\", {})
        all_of = rule.get(\"all_of\", [])
        any_of = rule.get(\"any_of\", [])
        none_of = rule.get(\"none_of\", [])
        if any(x not in raw_text for x in all_of):
            continue
        if any_of and not any(x in raw_text for x in any_of):
            continue
        if any(x in raw_text for x in none_of):
            continue
        return f.get(\"id\", \"default\"), _deep_merge(default_targets, f.get(\"targets\", {}))
    return \"default\", default_targets

def _extract_corp(raw_text: str, words: List[Dict[str, Any]], tcfg: Dict[str, Any]) -> Optional[str]:
    v = _regex_first(raw_text, tcfg.get(\"regex_patterns\", []))
    if v:
        return _post(v, tcfg.get(\"postprocess\", {}))
    gh = tcfg.get(\"geom_hints\", {})
    if gh:
        v = _find_near_anchor(words, gh.get(\"anchor_texts\", [\"様\"]), gh.get(\"x_tolerance\", 40), gh.get(\"y_up_range\", 60))
        if v:
            return _post(v, tcfg.get(\"postprocess\", {}))
    return None

def _kw_from_tables(pdf_bytes: bytes, tcfg: Dict[str, Any]) -> Optional[str]:
    if not tcfg.get(\"table\"):
        return None
    tb = tcfg[\"table\"]
    rows = tb.get(\"row_label_keywords\", [])
    headers = tb.get(\"header_kw_for_kw_col\", [])
    off = tb.get(\"fallback_col_offset_from_label\", 2)
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            try:
                tables = page.extract_tables()
            except Exception:
                tables = []
            for tbl in tables or []:
                norm = [[(c or \"\").strip() for c in row] for row in tbl if row]
                if not norm:
                    continue
                kw_col = None
                for j, cell in enumerate(norm[0]):
                    if any(h in cell for h in headers):
                        kw_col = j
                        break
                for row in norm:
                    label = row[0] if row else \"\"
                    if any(lbl in label for lbl in rows):
                        if kw_col is not None and kw_col < len(row):
                            m = re.search(r\"([0-9,]+\\.?[0-9]*)\", row[kw_col])
                            if m:
                                return m.group(1)
                        if len(row) > off:
                            m = re.search(r\"([0-9,]+\\.?[0-9]*)\", row[off])
                            if m:
                                return m.group(1)
    return None

def _kw_from_text(raw_text: str, tcfg: Dict[str, Any]) -> Optional[str]:
    rows = tcfg.get(\"table\", {}).get(\"row_label_keywords\", [])
    for ln in [x.strip() for x in raw_text.splitlines() if x.strip()]:
        if any(lbl in ln for lbl in rows):
            m = re.search(r\"([0-9,]+\\.?[0-9]*)\\s*(?:kW|ＫＷ|kw)\", ln)
            if m:
                return m.group(1)
    return _regex_first(raw_text, tcfg.get(\"regex_patterns\", []))

def _extract_pair(raw_text: str, words: List[Dict[str, Any]], targets: Dict[str, Any], pdf_bytes_for_tables: Optional[bytes]):
    out: Dict[str, Any] = {}
    corp_cfg = targets.get(\"法人名\", {})
    kw_cfg   = targets.get(\"契約電力\", {})
    out[\"法人名\"] = _extract_corp(raw_text, words, corp_cfg)
    kw_val = None
    if pdf_bytes_for_tables is not None:
        kw_val = _kw_from_tables(pdf_bytes_for_tables, kw_cfg)
    if not kw_val:
        kw_val = _kw_from_text(raw_text, kw_cfg)
    out[\"契約電力\"] = _post(kw_val, kw_cfg.get(\"postprocess\", {}))
    return out

def extract_data_text(pdf_bytes: bytes, config: Dict[str, Any]) -> Dict[str, Any]:
    raw_text, words = _extract_text_with_pdfplumber(pdf_bytes)
    form_id, targets = classify_form(raw_text, config)
    data = _extract_pair(raw_text, words, targets, pdf_bytes)
    data.update({\"raw_text\": raw_text, \"form_id\": form_id})
    return data

def extract_data_ocr(pdf_bytes: bytes, config: Dict[str, Any]) -> Dict[str, Any]:
    ocr_cfg = config.get(\"ocr\", {\"lang\": \"jpn+eng\", \"psm\": 6, \"oem\": 3, \"dpi\": 300})
    raw_text, words = _extract_text_with_ocr(pdf_bytes, ocr_cfg)
    form_id, targets = classify_form(raw_text, config)
    data = _extract_pair(raw_text, words, targets, None)
    data.update({\"raw_text\": raw_text, \"form_id\": form_id})
    return data

def _effective_targets(config: Dict[str, Any], form_id: Optional[str]) -> Dict[str, Any]:
    base = config.get(\"default_targets\", {})
    if not form_id or form_id == \"default\":
        return base
    for f in config.get(\"forms\", []):
        if f.get(\"id\") == form_id:
            return _deep_merge(base, f.get(\"targets\", {}))
    return base

def write_to_excel_with_mapping(data: Dict[str, Any], config: Dict[str, Any], template_path, form_id: Optional[str] = None) -> bytes:
    if hasattr(template_path, \"read\"):
        wb = load_workbook(filename=template_path)
    else:
        wb = load_workbook(filename=str(template_path))
    targets = _effective_targets(config, form_id)
    for key, tcfg in targets.items():
        sheet = tcfg.get(\"sheet\")
        cell  = tcfg.get(\"cell\")
        if sheet and cell and (key in data) and (data[key] is not None):
            try:
                wb[sheet][cell] = data[key]
            except KeyError:
                continue
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

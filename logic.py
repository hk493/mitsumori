from __future__ import annotations
import io, re, json, tempfile, cv2, numpy as np, unicodedata
from pathlib import Path
from typing import Dict, Tuple, Optional, List
import pandas as pd
import pdfplumber
from openpyxl import load_workbook, Workbook
from pdf2image import convert_from_bytes
from rapidocr_onnxruntime import RapidOCR

_RAPID_OCR = RapidOCR()


# ============ 設定ロード ============
def load_config(path: str = "config.json") -> dict:
    p = Path(path)
    if not p.exists():
        return {"ocr_preprocess": True, "ocr_dpi": 420, "ocr_keyword_box_margin": 40}
    with p.open("r", encoding="utf-8") as f:
        return json.load(f)


# ============ 日本語正規化＋簡体→日本語置換 ============
COMMON_FIXES = {
    "清求": "請求", "请求": "請求", "电": "電", "气": "気", "额": "額",
    "单": "単", "广": "広", "阳": "陽", "门": "門", "场": "場"
}
def normalize_japanese_text(s: str) -> str:
    t = unicodedata.normalize("NFKC", s)
    for k, v in COMMON_FIXES.items():
        t = t.replace(k, v)
    return re.sub(r"[ \u3000]{2,}", " ", t)


# ============ 法人名クリーンアップ ============
def _truncate_corporate_name(name: str, stop_words: List[str]) -> str:
    s = normalize_japanese_text(str(name).strip())
    s = re.split(r"(?:御中|様)", s)[0]
    if re.search(r"[0-9０-９]", s):
        return ""
    tails = ["株式会社","有限会社","医療法人","学校法人","合同会社","合名会社",
             "合資会社","病院","クリニック","大学","センター","組合","会社"]
    for t in sorted(tails, key=len, reverse=True):
        if t in s:
            return s[: s.find(t) + len(t)].strip()
    return s.strip()


def _num_to_clean_str(v) -> str:
    if v is None: return ""
    s = str(v).strip().replace(",", "")
    try:
        f = float(s)
        if abs(f - round(f)) < 1e-9:
            return str(int(round(f)))
        return f"{f}".rstrip("0").rstrip(".")
    except Exception:
        return s


# ============ テキストPDFか判定 ============
def _is_text_pdf(pdf_bytes: bytes) -> Tuple[bool, str]:
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            if not pdf.pages: return False, ""
            text = (pdf.pages[0].extract_text() or "").strip()
            return (len(text) > 20), text
    except Exception:
        return False, ""


# ============ OCR（RapidOCR + 前処理 + 領域再OCR） ============
def _ocr_text_and_words(pdf_bytes: bytes, cfg: dict) -> Tuple[str, List[dict]]:
    dpi = int(cfg.get("ocr_dpi", 420))
    pages = convert_from_bytes(pdf_bytes, dpi=dpi, first_page=1, last_page=3,
                               poppler_path=cfg.get("poppler_path"))
    preprocess = cfg.get("ocr_preprocess", True)
    margin = int(cfg.get("ocr_keyword_box_margin", 40))
    all_texts = []

    for img in pages:
        arr = np.array(img)
        if preprocess:
            gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
            gray = cv2.bilateralFilter(gray, 9, 75, 75)
            thr = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                        cv2.THRESH_BINARY,41,9)
            arr = cv2.cvtColor(thr, cv2.COLOR_GRAY2RGB)
        result, _ = _RAPID_OCR(arr)
        if not result: continue
        text = " ".join([r[1] for r in result])
        all_texts.append(text)

        # キーワード近傍を再OCR
        key_words = cfg.get("targets", {}).get("契約電力", {}).get("keywords_for_region_ocr", [])
        for r in result:
            if any(k in r[1] for k in key_words):
                box = r[0]  # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
                x_min = max(0, int(min(pt[0] for pt in box)))
                y_min = max(0, int(min(pt[1] for pt in box)))
                x_max = min(arr.shape[1], int(max(pt[0] for pt in box)) + margin)
                y_max = min(arr.shape[0], int(max(pt[1] for pt in box)) + margin)
                crop = arr[y_min:y_max, x_min:x_max]
                re_res, _ = _RAPID_OCR(crop)
                if re_res:
                    all_texts.append(" ".join([rr[1] for rr in re_res]))

    return "\n".join(all_texts).strip(), []


# ============ フィールド抽出 ============
def _extract_fields_from_text(text: str, cfg: dict) -> Dict[str, str]:
    out = {}
    text = normalize_japanese_text(text)

    # 法人名
    corp = None
    for pat in cfg.get("targets", {}).get("法人名", {}).get("regex", []):
        m = re.search(pat, text)
        if m:
            corp = m.group(1) if m.groups() else m.group(0)
            break
    if corp:
        corp = _truncate_corporate_name(corp, cfg.get("stop_keywords", []))
        if corp: out["法人名"] = corp

    # 契約電力
    kw = None
    for pat in cfg.get("targets", {}).get("契約電力", {}).get("regex", []):
        m = re.search(pat, text)
        if m:
            kw = _num_to_clean_str(m.group(1)); break
    if kw: out["契約電力"] = kw
    return out


# ============ 外向け ============
def process_pdf_bytes(pdf_bytes: bytes, cfg: dict) -> Tuple[Dict[str, str], str, bool]:
    is_text, text0 = _is_text_pdf(pdf_bytes)
    if is_text:
        text0 = normalize_japanese_text(text0)
        fields = _extract_fields_from_text(text0, cfg)
        return fields, text0, False
    text, _ = _ocr_text_and_words(pdf_bytes, cfg)
    fields = _extract_fields_from_text(text, cfg)
    return fields, text, True


# ============ Excel (最新月契約電力) ============
def _read_excel_to_df(xls_bytes: bytes, sheet_candidates: List[str]) -> pd.DataFrame:
    bio = io.BytesIO(xls_bytes)
    xl = pd.ExcelFile(bio)
    sheet = next((s for s in sheet_candidates if s in xl.sheet_names), xl.sheet_names[0])
    return xl.parse(sheet, header=None)


def process_excel_bytes(xls_bytes: bytes, cfg: dict) -> Tuple[Dict[str, str], pd.DataFrame]:
    excel_cfg = cfg.get("excel_input", {})
    df = _read_excel_to_df(xls_bytes, excel_cfg.get("sheet_candidates", ["高圧","Sheet1"]))
    H, W = df.shape
    result = {}
    stop_words = cfg.get("stop_keywords", [])

    # 法人名
    corp_raw = None
    for r in range(min(H,60)):
        for c in range(min(W,50)):
            v=df.iat[r,c]
            if v and any(k in str(v) for k in ["お客さま名","お客様名","会社名","法人名"]):
                corp_raw = df.iat[r,c+1] if c+1<W else None; break
        if corp_raw: break
    if corp_raw:
        corp=_truncate_corporate_name(corp_raw,stop_words)
        if corp: result["法人名"]=corp

    # 契約電力 最新月
    power_col=None; header_row=None
    for r in range(min(H,30)):
        for c in range(W):
            s=str(df.iat[r,c] or "")
            if "契約電力" in s and "kW" in s:
                power_col,header_row=c,r;break
        if power_col: break

    # 月列探索
    month_col=None
    if header_row is not None:
        for c in range(W):
            s=str(df.iat[header_row,c] or "")
            if "月" in s: month_col=c;break

    def parse_month(s):
        m=re.search(r"(20\d{2})\s*年\s*(\d{1,2})\s*月",s)
        return (int(m.group(1)),int(m.group(2))) if m else None

    latest_key=(-1,-1); latest_power=None
    if power_col is not None and header_row is not None:
        for r in range(header_row+1,H):
            ym=parse_month(str(df.iat[r,month_col])) if month_col else None
            if not ym: continue
            pv=df.iat[r,power_col]
            if pv and ym>latest_key:
                latest_key,latest_power=ym,_num_to_clean_str(pv)
    if latest_power: result["契約電力"]=latest_power
    return result,df


# ============ Excel書込 ============
def _ensure_template(path: Path, sheet: str):
    if not path.exists():
        wb=Workbook();ws=wb.active;ws.title=sheet
        ws["A1"]="法人名";ws["A12"]="契約電力";wb.save(path)

def write_to_excel(fields: Dict[str,str], cfg: dict, template_name="template_output.xlsx")->str:
    cell_map=cfg.get("excel_cell_map",{"sheet":"高圧","法人名":"B1","契約電力":"G12"})
    sheet=cell_map.get("sheet","高圧")
    path=Path(template_name);_ensure_template(path,sheet)
    wb=load_workbook(path)
    if sheet not in wb.sheetnames: wb.create_sheet(sheet)
    ws=wb[sheet]
    if "法人名" in fields: ws[cell_map["法人名"]]=fields["法人名"]
    if "契約電力" in fields: ws[cell_map["契約電力"]]=fields["契約電力"]
    tmp=tempfile.NamedTemporaryFile(delete=False,suffix=".xlsx")
    wb.save(tmp.name);tmp.close();return tmp.name

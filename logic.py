from __future__ import annotations
import io, re, json, tempfile
from pathlib import Path
from typing import Dict, Tuple, List
from pdf2image import convert_from_bytes
from paddleocr import PaddleOCR
import numpy as np
import cv2
import pandas as pd
from openpyxl import load_workbook

def load_config(path: str = "config.json") -> dict:
    import json
    with open(path, encoding="utf-8") as f:
        return json.load(f)

_OCR = None
def _get_ocr(lang: str = "japan", use_angle_cls: bool = True) -> PaddleOCR:
    global _OCR
    if _OCR is None:
        _OCR = PaddleOCR(
            lang=lang,
            use_angle_cls=use_angle_cls,
            show_log=False
        )
    return _OCR

def extract_month(text: str) -> int:
    # 「○月」の直前の数字を抽出
    match = re.search(r'(\d{1,2})月', text)
    if match:
        return int(match.group(1))
    return None

def extract_target_value(text: str) -> str:
    # キーワード近くの数値を抽出
    keywords = ["再エネ", "エネ", "燃料費", "燃料", "kWh"]
    for kw in keywords:
        # キーワードの近くにある数値（例：再エネ 12345）
        match = re.search(rf'{kw}[^\d]*([\d,]+)', text)
        if match:
            return match.group(1).replace(',', '')
    # kWhがついている数値（例：12345kWh）
    match = re.search(r'([\d,]+)\s*kWh', text)
    if match:
        return match.group(1).replace(',', '')
    return ""

def process_pdf_bytes(pdf_bytes: bytes, cfg: dict) -> Tuple[Dict, str]:
    dpi = cfg.get("ocr", {}).get("dpi", 200)
    max_pages = cfg.get("performance", {}).get("max_pages_per_pdf", 3)
    max_width = cfg.get("performance", {}).get("max_width", 1000)
    use_angle_cls = cfg.get("paddle", {}).get("use_angle_cls", False)
    lang = cfg.get("paddle", {}).get("lang", "japan")

    images = convert_from_bytes(pdf_bytes, dpi=dpi)
    texts = []
    fields = {}
    month = None
    value = ""
    for i, img in enumerate(images):
        if i >= max_pages:
            break
        img_np = np.array(img)
        h, w = img_np.shape[:2]
        if w > max_width:
            scale = max_width / w
            new_w = int(w * scale)
            new_h = int(h * scale)
            img_np = cv2.resize(img_np, (new_w, new_h), interpolation=cv2.INTER_AREA)
        if len(img_np.shape) == 3:
            img_np = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
        _, img_np = cv2.threshold(img_np, 180, 255, cv2.THRESH_BINARY)
        ocr = _get_ocr(lang, use_angle_cls)
        res = ocr.ocr(img_np, cls=use_angle_cls)
        page_text = "\n".join([t[1][0] for line in res for t in (line if isinstance(line, list) else [line])])
        texts.append(page_text)
        # 月判定
        if month is None:
            month = extract_month(page_text)
        # 金額抽出
        if not value:
            value = extract_target_value(page_text)
        # 既存フィールド抽出（法人名・契約電力）
        for k in cfg.get("targets", {}):
            for kw in cfg.get("excel_input", {}).get("label_keywords", {}).get(k, []):
                match = re.search(rf"{kw}[:：]?\s*([^\s\n]+)", page_text)
                if match:
                    fields[k] = match.group(1)
    # 月と値をfieldsに追加
    if month and value:
        fields[f"{month}月値"] = value
    return fields, "\n\n".join(texts)

def process_excel_bytes(excel_bytes: bytes, cfg: dict) -> Tuple[Dict, pd.DataFrame]:
    from io import BytesIO
    excel_io = BytesIO(excel_bytes)
    wb = load_workbook(excel_io, data_only=True)
    sheet_name = cfg.get("excel_cell_map", {}).get("sheet", wb.sheetnames[0])
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
    data = ws.values
    cols = next(data)
    df = pd.DataFrame(data, columns=cols)
    fields = {}
    # 簡易フィールド抽出例（法人名・契約電力）
    for k in cfg.get("targets", {}):
        for kw in cfg.get("excel_input", {}).get("label_keywords", {}).get(k, []):
            for col in df.columns:
                if kw in str(col):
                    fields[k] = df[col].iloc[0]
    return fields, df

def write_to_excel(list_of_fields: List[Dict], cfg: dict) -> str:
    template_path = Path("template_output.xlsx")
    if not template_path.exists():
        return ""
    wb = load_workbook(template_path)
    sheet_name = cfg.get("excel_cell_map", {}).get("sheet", wb.sheetnames[0])
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    # 既存項目（法人名・契約電力）は1行目に
    if list_of_fields:
        fields = list_of_fields[0]
        for key, cell in cfg.get("excel_cell_map", {}).items():
            if key in fields:
                ws[cell] = fields[key]  # ここでB1セルに法人名が入る

    # 月ごとの値をB21〜M21に代入
    month_cells = {
        1: "B21", 2: "C21", 3: "D21", 4: "E21", 5: "F21", 6: "G21",
        7: "H21", 8: "I21", 9: "J21", 10: "K21", 11: "L21", 12: "M21"
    }
    for fields in list_of_fields:
        for m in range(1, 13):
            key = f"{m}月値"
            if key in fields and month_cells.get(m):
                ws[month_cells[m]] = fields[key]
    out_path = "output_combined.xlsx"
    wb.save(out_path)
    return out_path

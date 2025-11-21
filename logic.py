from __future__ import annotations
import io, re, json, tempfile
from pathlib import Path
from typing import Dict, Tuple
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

def process_pdf_bytes(pdf_bytes: bytes, cfg: dict) -> Tuple[Dict, str]:
    dpi = cfg.get("ocr", {}).get("dpi", 200)
    max_pages = cfg.get("performance", {}).get("max_pages_per_pdf", 3)
    max_width = cfg.get("performance", {}).get("max_width", 1000)
    use_angle_cls = cfg.get("paddle", {}).get("use_angle_cls", False)
    lang = cfg.get("paddle", {}).get("lang", "japan")

    images = convert_from_bytes(pdf_bytes, dpi=dpi)
    texts = []
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
    return {}, "\n\n".join(texts)

def process_excel_bytes(excel_bytes: bytes, cfg: dict) -> Tuple[Dict, pd.DataFrame]:
    # ExcelファイルをDataFrameとしてプレビュー
    from io import BytesIO
    excel_io = BytesIO(excel_bytes)
    wb = load_workbook(excel_io, data_only=True)
    sheet_name = cfg.get("excel_cell_map", {}).get("sheet", wb.sheetnames[0])
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
    data = ws.values
    cols = next(data)
    df = pd.DataFrame(data, columns=cols)
    # fields抽出は必要に応じて追加
    return {}, df

def write_to_excel(fields: Dict, cfg: dict) -> str:
    # テンプレートExcelにfieldsを書き込む（必要に応じて実装）
    template_path = Path("template_output.xlsx")
    if not template_path.exists():
        return "テンプレートExcelがありません"
    wb = load_workbook(template_path)
    sheet_name = cfg.get("excel_cell_map", {}).get("sheet", wb.sheetnames[0])
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
    for key, cell in cfg.get("excel_cell_map", {}).items():
        if key in fields:
            ws[cell] = fields[key]
    out_path = "output.xlsx"
    wb.save(out_path)
    return out_path

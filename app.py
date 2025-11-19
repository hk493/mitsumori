import streamlit as st
import pandas as pd
from pathlib import Path
from logic import load_config, process_pdf_bytes, process_excel_bytes, write_to_excel

st.set_page_config(page_title="見積プロトタイプ（RapidOCR＋前処理）", layout="wide")
st.title("見積プロトタイプ｜PDF・Excel 明細 → テンプレExcel自動反映（RapidOCR版）")

cfg = load_config("config.json")

# セッション状態
if "extracted" not in st.session_state:
    st.session_state.extracted = None
if "raw_text" not in st.session_state:
    st.session_state.raw_text = ""
if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "used_ocr" not in st.session_state:
    st.session_state.used_ocr = False
if "excel_out" not in st.session_state:
    st.session_state.excel_out = None

# 3カラム構成
left, mid, right = st.columns([4, 1.5, 4])

with left:
    st.subheader("① 入力タイプを選択")
    mode = st.radio("アップロードする明細の形式を選択", ["PDF", "Excel明細"], horizontal=True)

    if mode == "PDF":
        pdf_file = st.file_uploader("PDFをアップロード（画像/テキストどちらでも可）", type=["pdf"])
        if pdf_file is not None:
            pdf_bytes = pdf_file.read()
            with st.spinner("PDF解析中…（RapidOCR + 前処理 + 正規化）"):
                fields, text, used_ocr = process_pdf_bytes(pdf_bytes, cfg)
                st.session_state.extracted = fields or {}
                st.session_state.raw_text = text or ""
                st.session_state.raw_df = None
                st.session_state.used_ocr = used_ocr

        st.markdown("### 生テキストプレビュー（PDF）")
        if st.session_state.raw_text:
            tag = "🟣 OCR使用（RapidOCR）" if st.session_state.used_ocr else "🟢 テキストPDF"
            st.caption(tag)
            st.text_area("抽出テキスト（正規化後）", st.session_state.raw_text, height=260)
        else:
            st.info("PDFをアップロードしてください。")

    else:
        xls_file = st.file_uploader("Excel明細をアップロード（.xlsx/.xlsm）", type=["xlsx", "xlsm", "xls"])
        if xls_file is not None:
            xls_bytes = xls_file.read()
            with st.spinner("Excel解析中…（セル・ラベル探索 + 最新月契約電力）"):
                fields, df = process_excel_bytes(xls_bytes, cfg)
                st.session_state.extracted = fields or {}
                st.session_state.raw_df = df
                st.session_state.raw_text = ""
                st.session_state.used_ocr = False

        st.markdown("### 生データプレビュー（Excel）")
        if st.session_state.raw_df is not None:
            st.dataframe(st.session_state.raw_df.head(50), use_container_width=True, height=260)
        else:
            st.info("Excel明細をアップロードしてください。")

with mid:
    st.subheader("② 実行")
    run_btn = st.button(
        "テンプレに書き込む",
        type="primary",
        use_container_width=True,
        disabled=st.session_state.extracted is None,
    )

    if run_btn:
        if st.session_state.extracted is None:
            st.warning("先にファイルをアップロードしてください。")
        else:
            out_path = write_to_excel(st.session_state.extracted, cfg, template_name="template_output.xlsx")
            st.session_state.excel_out = out_path
            st.success("Excelを書き出しました。右側からダウンロードできます。")

with right:
    st.subheader("③ 結果 & ダウンロード")

    if st.session_state.extracted is None:
        st.info("抽出結果はまだありません。")
    else:
        df = pd.DataFrame(list(st.session_state.extracted.items()), columns=["項目", "値"]).set_index("項目")
        if df.empty:
            st.warning("抽出項目は 0 件でした（テンプレには空のまま書き込み可能です）。")
        st.dataframe(df, use_container_width=True, height=180)

    if st.session_state.excel_out and Path(st.session_state.excel_out).exists():
        with open(st.session_state.excel_out, "rb") as f:
            data = f.read()
        st.download_button(
            label="結果Excelダウンロード",
            data=data,
            file_name="result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.info("処理するとここにダウンロードが出ます。")

st.divider()
st.caption(
    "テンプレは project 直下の `template_output.xlsx` を使用します。\n"
    "セル位置は config.json の `excel_cell_map` で調整可能。\n"
    "PDF は RapidOCR + 前処理 + 正規化 + 再OCR、Excel は最新月契約電力を自動取得。"
)

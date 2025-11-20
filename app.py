import streamlit as st
import pandas as pd
from pathlib import Path

from logic import load_config, process_pdf_bytes, process_excel_bytes, write_to_excel

st.set_page_config(page_title="見積プロトタイプ（PDF/Excel→テンプレ）", layout="wide")
st.title("見積プロトタイプ｜PDF/Excel 明細 → テンプレExcelへ自動反映")

cfg = load_config("config.json")

# セッション初期化
defaults = {
    "extracted": None,     # None: 未アップロード / {}: 抽出0件 / dict: 値あり
    "raw_text": "",        # PDFの全ページOCR結果
    "raw_df": None,        # Excelのプレビュー
    "excel_out": None,     # 書き出しパス
    "mode": "PDF（全ページOCR）"
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

left, mid, right = st.columns([4, 1.5, 4])

with left:
    st.subheader("① 入力タイプを選択")
    st.session_state.mode = st.radio(
        "アップロードする明細の形式",
        ["PDF（全ページOCR）", "Excel明細"],
        horizontal=True,
        index=0 if "PDF" in st.session_state.mode else 1
    )

    if "PDF" in st.session_state.mode:
        pdf_file = st.file_uploader("PDFをアップロード（全ページOCR／日本語）", type=["pdf"])
        if pdf_file is not None:
            pdf_bytes = pdf_file.read()
            with st.spinner("PaddleOCRで全ページOCR実行中…（前処理あり）"):
                fields, text = process_pdf_bytes(pdf_bytes, cfg)  # 全ページOCR
                st.session_state.extracted = fields if fields is not None else {}
                st.session_state.raw_text = text or ""
                st.session_state.raw_df = None
                st.session_state.excel_out = None

        st.markdown("### 生テキストプレビュー（全ページOCR結果）")
        if st.session_state.raw_text:
            st.text_area("抽出テキスト（全ページ）", st.session_state.raw_text, height=600)
        else:
            st.info("PDFをアップロードしてください。")

    else:
        xls_file = st.file_uploader("Excel明細をアップロード（.xlsx/.xlsm/.xls）", type=["xlsx", "xlsm", "xls"])
        if xls_file is not None:
            xls_bytes = xls_file.read()
            with st.spinner("Excel解析中…（最新月の契約電力を抽出）"):
                fields, df = process_excel_bytes(xls_bytes, cfg)
                st.session_state.extracted = fields if fields is not None else {}
                st.session_state.raw_df = df
                st.session_state.raw_text = ""
                st.session_state.excel_out = None

        st.markdown("### 生データプレビュー（Excel）")
        if st.session_state.raw_df is not None:
            st.dataframe(st.session_state.raw_df.head(80), use_container_width=True, height=320)
        else:
            st.info("Excel明細をアップロードしてください。")

with mid:
    st.subheader("② 実行")
    run_btn = st.button(
        "テンプレに書き込む",
        type="primary",
        use_container_width=True,
        disabled=(st.session_state.extracted is None),
    )

    if run_btn:
        if st.session_state.extracted is None:
            st.warning("先にファイルをアップロードしてください。")
        else:
            try:
                out_path = write_to_excel(
                    st.session_state.extracted,
                    cfg,
                    template_name="template_output.xlsx"  # 既存テンプレ必須
                )
                st.session_state.excel_out = out_path
                st.success("Excelを書き出しました。右側からダウンロードできます。")
            except FileNotFoundError as e:
                st.error(str(e))
            except Exception as e:
                st.exception(e)

with right:
    st.subheader("③ 結果 & ダウンロード")

    if st.session_state.extracted is None:
        st.info("抽出結果はまだありません。")
    else:
        df = pd.DataFrame(list(st.session_state.extracted.items()), columns=["項目", "値"]).set_index("項目")
        if df.empty:
            st.warning("抽出項目は 0 件でした（テンプレには空のまま書き込み可能です）。")
        st.dataframe(df, use_container_width=True, height=220)

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
    "テンプレはプロジェクト直下の `template_output.xlsx` を使用します（必須）。"
    "セル位置は config.json の `excel_cell_map` で調整できます。"
)

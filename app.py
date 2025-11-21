from pathlib import Path
import streamlit as st
from logic import load_config, process_pdf_bytes, process_excel_bytes, write_to_excel

st.set_page_config(page_title="見積プロトタイプ（PDF/Excel→テンプレ）", layout="wide")
st.title("見積プロトタイプ｜PDF/Excel 明細 → テンプレExcelへ自動反映")

cfg = load_config("config.json")

defaults = {
    "extracted": None,
    "raw_text": "",
    "raw_df": None,
    "excel_out": None,
    "mode": "PDF（全ページOCR）",
    "pdf_files": [],
    "excel_file": None,
    "processing": False,
    "output_files": [],
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
        pdf_files = st.file_uploader(
            "PDFをアップロード（複数選択可・一個ずつでもOK）",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_uploader"
        )
        if pdf_files:
            st.session_state.pdf_files = []
            for f in pdf_files:
                st.session_state.pdf_files.append({
                    "name": f.name,
                    "status": "未処理",
                    "data": {},
                    "text": "",
                    "bytes": f.read(),
                    "excel_path": ""
                })
    else:
        excel_file = st.file_uploader(
            "Excel明細ファイルをアップロード",
            type=["xlsx"],
            key="excel_uploader"
        )
        if excel_file:
            st.session_state.excel_file = excel_file

with mid:
    st.subheader("② 実行")
    if st.session_state.mode == "PDF（全ページOCR）":
        has_files = len(st.session_state.pdf_files) > 0
        run_btn = st.button(
            "OCR→Excelテンプレートに一括反映",
            type="primary",
            use_container_width=True,
            disabled=not has_files,
        )
        if run_btn and has_files:
            all_fields = []
            for idx, file_info in enumerate(st.session_state.pdf_files):
                st.session_state.pdf_files[idx]["status"] = "処理中"
                with st.spinner(f"🔄 {file_info['name']} をOCR実行中…"):
                    try:
                        fields, text = process_pdf_bytes(file_info["bytes"], cfg)
                        st.session_state.pdf_files[idx]["status"] = "完了"
                        st.session_state.pdf_files[idx]["data"] = fields if fields else {}
                        st.session_state.pdf_files[idx]["text"] = text or ""
                        all_fields.append(fields)
                        st.success(f"✅ {file_info['name']} の処理が完了しました")
                    except Exception as e:
                        st.session_state.pdf_files[idx]["status"] = "エラー"
                        st.error(f"❌ {file_info['name']} の処理中にエラーが発生しました: {str(e)}")
            # まとめてExcelに代入
            excel_path = write_to_excel(all_fields, cfg)
            st.session_state.output_files = excel_path

with right:
    st.subheader("③ 結果プレビュー・ダウンロード")
    if st.session_state.mode == "PDF（全ページOCR）":
        if st.session_state.output_files and Path(st.session_state.output_files).exists():
            with open(st.session_state.output_files, "rb") as f:
                st.download_button(
                    label="まとめてExcelダウンロード",
                    data=f.read(),
                    file_name="output_combined.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        # OCRテキストプレビューは個別に表示
        if st.session_state.pdf_files:
            for file_info in st.session_state.pdf_files:
                st.write(f"**{file_info['name']}** - {file_info['status']}")
                if file_info["status"] == "完了":
                    st.text_area("OCRテキスト", file_info["text"], height=150)
                elif file_info["status"] == "エラー":
                    st.write("エラーが発生しました。")
    else:
        if st.session_state.raw_df is not None:
            st.dataframe(st.session_state.raw_df)
        if st.session_state.excel_out and Path(st.session_state.excel_out).exists():
            with open(st.session_state.excel_out, "rb") as f:
                st.download_button(
                    label="Excelダウンロード",
                    data=f.read(),
                    file_name="output_excel.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.divider()
st.caption(
    "テンプレはプロジェクト直下の `template_output.xlsx` を使用します（必須）。"
    "セル位置は config.json の `excel_cell_map` で調整できます。"
)

import io, json, streamlit as st
from logic import load_config, extract_data_text, extract_data_ocr, write_to_excel_with_mapping

st.set_page_config(page_title=\"estimation-proto (手動切替)\", layout=\"wide\")
st.title(\"estimation-proto（手動：Python解析 / OCR解析）\")

col_left, col_center, col_right = st.columns([3, 2, 2])

for k, v in {\"raw_text\":\"\", \"extracted\":{}, \"result_bytes\":None, \"result_name\":None, \"config\":None, \"form_id\":None}.items():
    if k not in st.session_state:
        st.session_state[k] = v

with col_left:
    st.subheader(\"1) アップロード\")
    template_file = st.file_uploader(\"Excelテンプレ（任意: 未指定なら template_output.xlsx を使用）\", type=[\"xlsx\"])
    cfg_file      = st.file_uploader(\"config.json（任意）\", type=[\"json\"])
    if cfg_file:
        try:
            st.session_state[\"config\"] = json.load(cfg_file)
            st.caption(\"📄 アップロードした config.json を使用します。\")
        except Exception as e:
            st.error(f\"config.json 読み込み失敗: {e}\")

    text_pdf = st.file_uploader(\"テキストPDF（文字埋込み）\", type=[\"pdf\"])
    ocr_pdf  = st.file_uploader(\"スキャンPDF（OCR用）\", type=[\"pdf\"])

    if st.session_state[\"raw_text\"]:
        with st.expander(\"📜 抽出テキスト（先頭1万文字）\"):
            st.text(st.session_state[\"raw_text\"][:10000])
        st.markdown(\"### 抽出結果プレビュー\")

        def pill(label, value):
            if value in (None, \"\"):
                st.markdown(\"<div style='padding:8px;border-radius:8px;background:#ffe6e6;color:#a10000'>❌ {}: 未検出</div>\".format(label), unsafe_allow_html=True)
            else:
                st.markdown(\"<div style='padding:8px;border-radius:8px;background:#e6ffef;color:#0b6b35'>✅ {}: {}</div>\".format(label, value), unsafe_allow_html=True)

        pill(\"法人名\",       st.session_state[\"extracted\"].get(\"法人名\"))
        pill(\"契約電力(kW)\", st.session_state[\"extracted\"].get(\"契約電力\"))
        if st.session_state.get(\"form_id\"):
            st.info(f\"判別フォームID: {st.session_state['form_id']}\")
with col_center:
    st.subheader(\"2) 解析ボタン（手動）\")
    if st.button(\"テキストを解析（Python）\", use_container_width=True):
        if not text_pdf:
            st.error(\"テキストPDFを選んでください。\")
        else:
            try:
                cfg = st.session_state[\"config\"] or load_config(\"config.json\")
                result = extract_data_text(text_pdf.read(), cfg)
                st.session_state[\"raw_text\"]  = result.get(\"raw_text\", \"\")
                st.session_state[\"extracted\"] = {k: result.get(k) for k in [\"法人名\", \"契約電力\"]}
                st.session_state[\"form_id\"]   = result.get(\"form_id\")
                st.success(\"テキストPDFの解析が完了しました。\")
            except Exception as e:
                st.error(f\"解析失敗: {e}\")

    if st.button(\"OCRで解析（Tesseract）\", use_container_width=True):
        if not ocr_pdf:
            st.error(\"OCR用PDFを選んでください。\")
        else:
            try:
                cfg = st.session_state[\"config\"] or load_config(\"config.json\")
                result = extract_data_ocr(ocr_pdf.read(), cfg)
                st.session_state[\"raw_text\"]  = result.get(\"raw_text\", \"\")
                st.session_state[\"extracted\"] = {k: result.get(k) for k in [\"法人名\", \"契約電力\"]}
                st.session_state[\"form_id\"]   = result.get(\"form_id\")
                st.success(\"OCR解析が完了しました。\")
            except Exception as e:
                st.error(f\"OCR解析失敗: {e}\")

    st.markdown(\"---\")
    if st.button(\"テンプレに書き込み → Excel生成\", type=\"primary\", use_container_width=True):
        if not st.session_state[\"extracted\"]:
            st.error(\"先に解析を実行してください。\")
        else:
            try:
                cfg = st.session_state[\"config\"] or load_config(\"config.json\")
                tpath = io.BytesIO(template_file.getvalue()) if template_file else \"template_output.xlsx\"
                result_bytes = write_to_excel_with_mapping(
                    data=st.session_state[\"extracted\"],
                    config=cfg,
                    template_path=tpath,
                    form_id=st.session_state.get(\"form_id\"),
                )
                st.session_state[\"result_bytes\"] = result_bytes
                st.session_state[\"result_name\"]  = \"output_estimation.xlsx\"
                st.success(\"Excelを生成しました。右カラムからダウンロードできます。\")
            except Exception as e:
                st.error(f\"Excel生成失敗: {e}\")
with col_right:
    st.subheader(\"3) ダウンロード\")
    if st.session_state[\"result_bytes\"] is not None:
        st.download_button(
            \"結果Excelをダウンロード\",
            st.session_state[\"result_bytes\"],
            file_name=st.session_state[\"result_name\"],
            mime=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\",
            use_container_width=True,
        )
    else:
        st.caption(\"まだ生成物がありません。中央のボタンでExcelを作成してください。\")

import tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
from logic import load_config, extract_data, write_to_excel_with_mapping

st.set_page_config(page_title="見積プロトタイプ（文字列そのまま出力）", layout="wide")
st.title("PDF→Excel（法人名/契約電力）プロトタイプ")

if "result" not in st.session_state:
    st.session_state.result = None
if "download_path" not in st.session_state:
    st.session_state.download_path = None

left, mid, right = st.columns([4, 2, 4])

with left:
    st.subheader("① テンプレ & PDF アップロード")
    tpl_upload = st.file_uploader(
        "Excelテンプレ（未選択なら ./template_output.xlsx を使用）", type=["xlsx"]
    )
    pdf_file = st.file_uploader("PDF（1社1ページ想定・テキストベース）", type=["pdf"])

    cfg = load_config("config.json")

    if pdf_file is not None:
        res = extract_data(pdf_file.read(), cfg)
        st.session_state.result = res

        st.markdown("### 生テキスト（抜粋）")
        st.text_area("テキスト", res.get("raw_text", "")[:8000], height=260)

        st.markdown("### 抽出結果（未取得は赤）")
        order = list(cfg["targets"].keys())
        rows = [[k, res.get(k)] for k in order]
        df = pd.DataFrame(rows, columns=["項目", "値"]).set_index("項目")

        def hi(s: pd.Series):
            return [
                "background-color:#ffe5e5; color:#b00020; font-weight:bold;"
                if (v is None or v == "")
                else ""
                for v in s.values
            ]

        st.dataframe(df.style.apply(hi, axis=1), use_container_width=True, height=150)
    else:
        st.info("PDFをアップロードしてください。")

with mid:
    st.subheader("② 見積生成")
    st.write("抽出文字列をテンプレ指定セルへ“そのまま”書き込みます。")

    run = st.button(
        "処理を実行（Excel作成）",
        type="primary",
        use_container_width=True,
        disabled=not bool(st.session_state.result),
    )

    if run:
        cfg = load_config("config.json")
        data_to_fill = {k: st.session_state.result.get(k) for k in cfg["targets"].keys()}

        if tpl_upload is not None:
            t = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            t.write(tpl_upload.read())
            t.flush()
            t.close()
            tpl_path = t.name
        else:
            tpl_path = "template_output.xlsx"

        out_path = write_to_excel_with_mapping(data_to_fill, cfg, tpl_path)
        st.session_state.download_path = out_path
        st.success("Excelを作成しました。右側からダウンロードできます。")

with right:
    st.subheader("③ ダウンロード")
    if st.session_state.download_path and Path(st.session_state.download_path).exists():
        with open(st.session_state.download_path, "rb") as f:
            data = f.read()
        corp = (st.session_state.result or {}).get("法人名") or "result"
        st.download_button(
            "結果Excelをダウンロード",
            data=data,
            file_name=f"{corp}_見積.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.info("処理後にここにダウンロードが表示されます。")

st.divider()
st.caption("法人名は『病院／医院／クリニック／会社』で切り、契約電力は主契約の右側→基本料金×kW列→表→正規表現の順で抽出。")

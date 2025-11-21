import streamlit as st
import pandas as pd
from pathlib import Path
import time

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
    "mode": "PDF（全ページOCR）",
    "pdf_files": [],       # 複数PDFファイルの情報 [{"name": str, "status": str, "data": dict, "text": str}]
    "processing": False,   # OCR処理中フラグ
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
        # ファイルアップローダー
        pdf_file = st.file_uploader(
            "PDFをアップロード（全ページOCR／日本語）", 
            type=["pdf"],
            key="pdf_uploader"
        )
        
        # 新しいファイルがアップロードされた場合の処理
        if pdf_file is not None:
            # 既に同じファイル名が処理済みまたは処理中かチェック
            existing_names = [f["name"] for f in st.session_state.pdf_files]
            
            if pdf_file.name not in existing_names:
                # 処理中のファイルがあるかチェック
                has_processing = any(f["status"] == "処理中" for f in st.session_state.pdf_files)
                
                if has_processing:
                    # 処理中のファイルがある場合は待機状態で追加
                    st.session_state.pdf_files.append({
                        "name": pdf_file.name,
                        "status": "待機中",
                        "data": {},
                        "text": "",
                        "bytes": pdf_file.read()
                    })
                    st.info(f"📄 {pdf_file.name} を追加しました。前のファイルの処理完了後に自動的に処理されます。")
                else:
                    # 処理中のファイルがない場合は即座に処理開始
                    pdf_bytes = pdf_file.read()
                    st.session_state.pdf_files.append({
                        "name": pdf_file.name,
                        "status": "処理中",
                        "data": {},
                        "text": "",
                        "bytes": pdf_bytes
                    })
                    st.session_state.processing = True
                    
                    with st.spinner(f"🔄 {pdf_file.name} をOCR実行中…"):
                        try:
                            fields, text = process_pdf_bytes(pdf_bytes, cfg)
                            st.session_state.pdf_files[-1]["status"] = "完了"
                            st.session_state.pdf_files[-1]["data"] = fields if fields else {}
                            st.session_state.pdf_files[-1]["text"] = text or ""
                            st.success(f"✅ {pdf_file.name} の処理が完了しました")
                        except Exception as e:
                            st.session_state.pdf_files[-1]["status"] = "エラー"
                            st.error(f"❌ {pdf_file.name} の処理中にエラーが発生しました: {str(e)}")
                    
                    st.session_state.processing = False
                    st.rerun()
        
        # 待機中のファイルがあれば自動的に処理
        waiting_files = [f for f in st.session_state.pdf_files if f["status"] == "待機中"]
        processing_files = [f for f in st.session_state.pdf_files if f["status"] == "処理中"]
        
        if waiting_files and not processing_files:
            # 待機中の最初のファイルを処理
            next_file = waiting_files[0]
            next_file["status"] = "処理中"
            st.session_state.processing = True
            
            with st.spinner(f"🔄 {next_file['name']} をOCR実行中…"):
                try:
                    fields, text = process_pdf_bytes(next_file["bytes"], cfg)
                    next_file["status"] = "完了"
                    next_file["data"] = fields if fields else {}
                    next_file["text"] = text or ""
                    st.success(f"✅ {next_file['name']} の処理が完了しました")
                except Exception as e:
                    next_file["status"] = "エラー"
                    st.error(f"❌ {next_file['name']} の処理中にエラーが発生しました: {str(e)}")
            
            st.session_state.processing = False
            st.rerun()
        
        # ファイル一覧の表示
        if st.session_state.pdf_files:
            st.markdown("### 📁 処理ファイル一覧")
            
            # 全体の進捗状況
            total_files = len(st.session_state.pdf_files)
            completed_count = len([f for f in st.session_state.pdf_files if f["status"] == "完了"])
            error_count = len([f for f in st.session_state.pdf_files if f["status"] == "エラー"])
            
            # プログレスバー
            progress = completed_count / total_files if total_files > 0 else 0
            st.progress(progress, text=f"進行状況: {completed_count}/{total_files} ファイル完了")
            
            # 各ファイルの状態
            for idx, file_info in enumerate(st.session_state.pdf_files, 1):
                status_icon = {
                    "待機中": "⏳",
                    "処理中": "🔄",
                    "完了": "✅",
                    "エラー": "❌"
                }.get(file_info["status"], "❓")
                
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.text(f"{idx}. {status_icon} {file_info['name']}")
                with col2:
                    st.text(f"{file_info['status']}")
        
        # 全ファイルのテキストプレビュー
        st.markdown("### 生テキストプレビュー（全ページOCR結果）")
        completed_files = [f for f in st.session_state.pdf_files if f["status"] == "完了"]
        if completed_files:
            all_text = "\n\n=== ファイル区切り ===\n\n".join([
                f"【{f['name']}】\n{f['text']}" for f in completed_files
            ])
            st.text_area("抽出テキスト（全ページ）", all_text, height=600)
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
    
    # ボタンの有効化条件を判定
    if "PDF" in st.session_state.mode:
        # PDF: すべてのファイルが完了している かつ 1つ以上のファイルがある
        all_completed = all(f["status"] == "完了" for f in st.session_state.pdf_files)
        has_files = len(st.session_state.pdf_files) > 0
        is_disabled = not (all_completed and has_files)
    else:
        # Excel: 従来通り
        is_disabled = st.session_state.extracted is None
    
    run_btn = st.button(
        "テンプレに書き込む",
        type="primary",
        use_container_width=True,
        disabled=is_disabled,
    )

    if run_btn:
        if "PDF" in st.session_state.mode:
            # 全PDFファイルのデータを統合
            all_data = {}
            for file_info in st.session_state.pdf_files:
                if file_info["status"] == "完了":
                    # 既存のデータに追加（後のファイルで上書き）
                    all_data.update(file_info["data"])
            
            if not all_data:
                st.warning("抽出されたデータがありません。")
            else:
                try:
                    out_path = write_to_excel(
                        all_data,
                        cfg,
                        template_name="template_output.xlsx"
                    )
                    st.session_state.excel_out = out_path
                    st.session_state.extracted = all_data
                    st.success("Excelを書き出しました。右側からダウンロードできます。")
                except FileNotFoundError as e:
                    st.error(str(e))
                except Exception as e:
                    st.exception(e)
        else:
            # Excel処理（従来通り）
            if st.session_state.extracted is None:
                st.warning("先にファイルをアップロードしてください。")
            else:
                try:
                    out_path = write_to_excel(
                        st.session_state.extracted,
                        cfg,
                        template_name="template_output.xlsx"
                    )
                    st.session_state.excel_out = out_path
                    st.success("Excelを書き出しました。右側からダウンロードできます。")
                except FileNotFoundError as e:
                    st.error(str(e))
                except Exception as e:
                    st.exception(e)

with right:
    st.subheader("③ 結果 & ダウンロード")

    # PDF複数ファイル対応の表示
    if "PDF" in st.session_state.mode and st.session_state.pdf_files:
        completed_files = [f for f in st.session_state.pdf_files if f["status"] == "完了"]
        
        if completed_files:
            # OCR完了ファイル一覧
            st.markdown("### ✅ OCR完了ファイル")
            for idx, file_info in enumerate(completed_files, 1):
                with st.expander(f"{idx}. {file_info['name']}", expanded=False):
                    if file_info["data"]:
                        df = pd.DataFrame(
                            list(file_info["data"].items()), 
                            columns=["項目", "値"]
                        ).set_index("項目")
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.info("抽出データなし")
            
            st.divider()
            
            # 統合データの表示
            all_data = {}
            for f in completed_files:
                all_data.update(f["data"])
            
            if all_data:
                st.markdown("### 📊 統合結果")
                df = pd.DataFrame(
                    list(all_data.items()), 
                    columns=["項目", "値"]
                ).set_index("項目")
                st.dataframe(df, use_container_width=True, height=150)
        else:
            st.info("OCR処理が完了したファイルはまだありません。")
    elif st.session_state.extracted is None:
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

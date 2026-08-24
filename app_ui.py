import streamlit as st
from openai import OpenAI
import os
from dotenv import load_dotenv
from pypdf import PdfReader
from docx import Document
import pandas as pd
import base64
import tempfile
import shutil
from openpyxl import load_workbook

# ★同じフォルダにある「.env」ファイルから設定を自動で読み込む
# load_dotenv(override=True)
load_dotenv("D:\\OpenAI_Test\\func.env")

# ★環境変数から安全にURLとキーを取得（コード上に直接書かない！）
#endpoint = os.getenv("AZURE_AI_ENDPOINT")
#api_key = os.getenv("AZURE_AI_KEY")
endpoint = st.secrets["AZURE_AI_ENDPOINT"]
api_key  = st.secrets["AZURE_AI_KEY"]

# 画面のタイトル設定
st.title("統合資料・画像認識対応 AIアシスタント (RAG)")
st.caption("PDF、Word、Excel、および画像ファイル（図）をまとめて同時に読み込んで質問できます")

client = OpenAI(base_url=endpoint, api_key=api_key)


# ★新機能: Excelの「グラフ」機能で作成したネイティブチャートを画像化する
# ネイティブチャートはxlsx内に画像として保存されておらず描画情報（XML）しか入っていないため、
# openpyxlでは取り出せない。ローカルにインストール済みのExcelをCOM経由で自動操作し、
# 各グラフをPNGとしてExportしてから読み込む。
# 前提: Windows環境 かつ Microsoft Excelがインストール済み かつ `pip install pywin32` 済みであること。
def extract_excel_charts_via_com(file_bytes, filename):
    images = []

    try:
        import win32com.client as win32
    except ImportError:
        st.warning(
            "ネイティブグラフの抽出には pywin32 が必要です。"
            "コマンドプロンプトで `pip install pywin32` を実行してください。"
        )
        return images

    xlChart = -4109  # Excel定数 XlSheetType.xlChart（グラフシート判定用）

    tmp_dir = tempfile.mkdtemp(prefix="xlsx_chart_")
    tmp_xlsx_path = os.path.join(tmp_dir, filename)
    with open(tmp_xlsx_path, "wb") as f:
        f.write(file_bytes)

    excel = None
    wb = None
    try:
        # 既存で起動中のExcelに影響を与えないよう、新規プロセスとして起動する
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(tmp_xlsx_path, ReadOnly=True, UpdateLinks=False)

        chart_index = 0
        for sheet in wb.Sheets:
            # ケース1: 通常のワークシートに埋め込まれたグラフ（ChartObjects）
            if sheet.Type != xlChart:
                for chart_obj in sheet.ChartObjects():
                    chart_index += 1
                    img_path = os.path.join(tmp_dir, f"chart_{chart_index}.png")
                    chart_obj.Chart.Export(img_path, "PNG")
                    with open(img_path, "rb") as img_f:
                        images.append({
                            "data": base64.b64encode(img_f.read()).decode("utf-8"),
                            "mime": "image/png",
                        })
            # ケース2: シート全体がグラフになっている「グラフシート」
            else:
                chart_index += 1
                img_path = os.path.join(tmp_dir, f"chart_{chart_index}.png")
                sheet.Export(img_path, "PNG")
                with open(img_path, "rb") as img_f:
                    images.append({
                        "data": base64.b64encode(img_f.read()).decode("utf-8"),
                        "mime": "image/png",
                    })
    except Exception as e:
        st.warning(f"「{filename}」のネイティブグラフ抽出中にエラーが発生しました: {e}")
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        shutil.rmtree(tmp_dir, ignore_errors=True)

    return images


# アップロードされた画像を保持するリスト
if "extracted_images" not in st.session_state:
    st.session_state.extracted_images = []

with st.sidebar:
    st.header("資料・図のアップロード")
    # ★ 拡張子に「png」「jpg」「jpeg」を追加して、図の画像を直接ドロップできるように解放
    uploaded_files = st.file_uploader(
        "ファイルを選択してください（複数可）", 
        type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"],
        accept_multiple_files=True
    )
    
    combined_text = ""
    
    if uploaded_files:
        st.session_state.extracted_images = []
        
        for uploaded_file in uploaded_files:
            file_ext = uploaded_file.name.split(".")[-1].lower()
            
            # ① PDFの読み込み
            if file_ext == "pdf":
                reader = PdfReader(uploaded_file)
                file_text = ""
                for page in reader.pages:
                    file_text += page.extract_text() or ""
                if file_text.strip():
                    combined_text += f"\n\n--- ファイル名: {uploaded_file.name} ---\n{file_text}"
                    st.success(f"「{uploaded_file.name}」のテキスト読み込み成功")
                    
            # ② Wordの読み込み
            elif file_ext == "docx":
                doc = Document(uploaded_file)
                file_text = ""
                for paragraph in doc.paragraphs:
                    file_text += paragraph.text + "\n"
                if file_text.strip():
                    combined_text += f"\n\n--- ファイル名: {uploaded_file.name} ---\n{file_text}"
                    st.success(f"「{uploaded_file.name}」のテキスト読み込み成功")
                    
            # ③ Excelの読み込み（セルのテキスト＋埋め込み画像の両方を抽出）
            elif file_ext == "xlsx":
                excel_file = pd.ExcelFile(uploaded_file)
                file_text = ""
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                    file_text += f"\n【シート名: {sheet_name}】\n"
                    file_text += df.to_string(index=False) + "\n"
                if file_text.strip():
                    combined_text += f"\n\n--- ファイル名: {uploaded_file.name} ---\n{file_text}"
                    st.success(f"「{uploaded_file.name}」のテキスト読み込み成功")

                # ★新機能: シートに貼り付けられた画像（スクリーンショット・図など）を抽出
                # pandasはセルの値しか読まないため、貼り付けられた画像は無視されていた。
                # openpyxlでワークブックを開き、各シートに埋め込まれた画像を直接取り出してAIに渡す。
                uploaded_file.seek(0)  # pd.ExcelFileで進んだ読み込み位置を先頭に戻す
                try:
                    wb = load_workbook(uploaded_file)
                    image_count = 0
                    for ws in wb.worksheets:
                        for image in getattr(ws, "_images", []):
                            img_bytes = image._data()
                            img_str = base64.b64encode(img_bytes).decode('utf-8')
                            img_format = (getattr(image, "format", None) or "png").lower()
                            st.session_state.extracted_images.append({
                                "data": img_str,
                                "mime": f"image/{img_format}",
                            })
                            image_count += 1
                    if image_count > 0:
                        st.success(f"「{uploaded_file.name}」内の画像 {image_count} 件をAIの「目」として抽出しました")
                except Exception as e:
                    st.warning(
                        f"「{uploaded_file.name}」の画像抽出中にエラーが発生しました（テキストは読み込み済みです）: {e}"
                    )

                # ★新機能: Excelの「グラフ」機能で作成したネイティブチャートをExcel COM経由で画像化
                # ※Windows環境＋Excelインストール済み＋pywin32導入済みの場合のみ有効
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                chart_images = extract_excel_charts_via_com(file_bytes, uploaded_file.name)
                if chart_images:
                    st.session_state.extracted_images.extend(chart_images)
                    st.success(
                        f"「{uploaded_file.name}」内のグラフ {len(chart_images)} 件をAIの「目」として抽出しました"
                    )

            # ④ ★新機能: 画像ファイル（図）が直接アップロードされた場合の処理
            elif file_ext in ["png", "jpg", "jpeg"]:
                img_data = uploaded_file.read()
                img_str = base64.b64encode(img_data).decode('utf-8')
                img_mime = "image/jpeg" if file_ext in ["jpg", "jpeg"] else "image/png"
                st.session_state.extracted_images.append({
                    "data": img_str,
                    "mime": img_mime,
                })
                st.success(f"図・画像「{uploaded_file.name}」を正常に認識しました")

# チャット履歴を保持する仕組み
if "messages" not in st.session_state:
    st.session_state.messages = []

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

# ユーザーからの入力欄
if prompt := st.chat_input("アップロードしたすべての資料や図について質問してください..."):
    with st.chat_message("user"):
        st.markdown(prompt)
    st.session_state.messages.append({"role": "user", "content": prompt})

    with st.chat_message("assistant"):
        system_instruction = "あなたは優秀なアシスタントです。提供された【参考資料のテキスト情報】、および添付された「画像（図）」の内容を人間の目のように厳密に確認し、ユーザーの質問に正確に答えてください。画像に描かれているネットワーク構成、サブネット名、IPアドレス等の文字情報もすべて読み取って回答に反映してください。"
        if combined_text:
            system_instruction += f"\n\n【参考資料のテキスト情報】\n{combined_text}"

        # AIに送るメッセージを画像対応のマルチモーダル形式に変換
        content_list = [{"type": "text", "text": prompt}]
        
        # アップロードされたすべての画像（直接添付＋Excel等から抽出したもの）をAIの「目」として添付する
        for img_info in st.session_state.extracted_images:
            content_list.append({
                "type": "image_url",
                "image_url": {"url": f"data:{img_info['mime']};base64,{img_info['data']}"}
            })

        response = client.chat.completions.create(
            model="gpt-5.4-mini-1", 
            messages=[
                {"role": "system", "content": system_instruction},
                {"role": "user", "content": content_list}
            ]
        )
        answer = response.choices[0].message.content
        st.markdown(answer)
    
    st.session_state.messages.append({"role": "assistant", "content": answer})

import streamlit as st
from openai import OpenAI  
import os  
from dotenv import load_dotenv  
from pypdf import PdfReader  
from docx import Document  
import pandas as pd  
import base64

# ★同じフォルダにある「.env」ファイルから設定を自動で読み込む
# load_dotenv(override=True)
load_dotenv("D:\\OpenAI_Test\\func.env")

# ★環境変数から安全にURLとキーを取得（コード上に直接書かない！）
endpoint = os.getenv("AZURE_AI_ENDPOINT")
api_key = os.getenv("AZURE_AI_KEY")

# 画面のタイトル設定
st.title("🤖 統合資料・画像認識対応 AIアシスタント (RAG)")
st.caption("PDF、Word、Excel、および画像ファイル（図）をまとめて同時に読み込んで質問できます")

client = OpenAI(base_url=endpoint, api_key=api_key)

# アップロードされた画像を保持するリスト
if "extracted_images" not in st.session_state:
    st.session_state.extracted_images = []

with st.sidebar:
    st.header("📄 資料・図のアップロード")
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
                    
            # ③ Excelの読み込み
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
            
            # ④ ★新機能: 画像ファイル（図）が直接アップロードされた場合の処理
            elif file_ext in ["png", "jpg", "jpeg"]:
                img_data = uploaded_file.read()
                img_str = base64.b64encode(img_data).decode('utf-8')
                st.session_state.extracted_images.append(img_str)
                st.success(f"📸 図・画像「{uploaded_file.name}」を正常に認識しました")

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
        
        # アップロードされたすべての画像をAIの「目」として添付する
        for img_base64 in st.session_state.extracted_images:
            content_list.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{img_base64}"}
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

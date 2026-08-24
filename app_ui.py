import streamlit as st
from openai import OpenAI
import os
from dotenv import load_dotenv
from pypdf import PdfReader
from docx import Document
import pandas as pd
import base64
import sys
import platform
import subprocess
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


# ★新機能: Excelの「グラフ」「図形（オートシェイプ／SmartArt／グループ化した図形）」
# 「挿入した埋め込みオブジェクト（Visio図面など）」をまとめて画像化する。
# 構成図は多くの場合、複数の図形（四角・線・矢印・テキストボックス等）を組み合わせて描かれているか、
# 「挿入」→「オブジェクト」で貼り込んだ埋め込みオブジェクトになっている。これらはopenpyxlの
# ws._images（貼り付け画像専用）や単純なChartObjects列挙だけでは取り出せないため、
# ローカルにインストール済みのExcelをCOM（pywin32）経由で自動操作し、
# 「シート上の図形をまとめて選択してコピー→仮のグラフに貼り付けてPNGとしてExport」
# というExcel VBAでもよく使われる手法でスクリーンショット化する。
# 前提: Windows環境 かつ Microsoft Excelがインストール済み かつ `pip install pywin32` 済みであること。
def extract_excel_diagrams_via_com(file_bytes, filename):
    images = []

    try:
        import win32com.client as win32
    except ImportError:
        # ★診断用: このStreamlitプロセスが実際にどのpython.exeで動いているかを表示する。
        # 「別のターミナル/仮想環境にpywin32を入れたのに反映されない」場合は、
        # ここに表示されたパスと、pip installを実行したPythonのパスが違うことが多い。
        st.warning(
            "図形・オブジェクトの抽出には pywin32 が必要です。\n\n"
            f"現在このStreamlitアプリを実行しているPython: `{sys.executable}`\n\n"
            "このPythonに対して以下を実行してから、"
            "**streamlit runしているターミナルを完全に一度停止（Ctrl+C）して再起動**してください"
            "（ブラウザの「Rerun」だけではプロセスが再起動されず反映されません）。\n\n"
            f"```\n\"{sys.executable}\" -m pip install pywin32\n"
            f"\"{os.path.dirname(sys.executable)}\\python.exe\" "
            f"\"{os.path.dirname(sys.executable)}\\Scripts\\pywin32_postinstall.py\" -install\n```"
        )
        return images

    import time

    xlChart = -4109  # Excel定数 XlSheetType.xlChart（グラフシート判定用）
    msoPicture = 13  # 単純な貼り付け画像はopenpyxl側で既に抽出済みのため対象外にする

    tmp_dir = tempfile.mkdtemp(prefix="xlsx_diagram_")
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

        img_index = 0
        for sheet in wb.Sheets:
            # ケース1: シート全体がグラフになっている「グラフシート」
            if sheet.Type == xlChart:
                img_index += 1
                img_path = os.path.join(tmp_dir, f"diagram_{img_index}.png")
                sheet.Export(img_path, "PNG")
                with open(img_path, "rb") as img_f:
                    images.append({
                        "data": base64.b64encode(img_f.read()).decode("utf-8"),
                        "mime": "image/png",
                    })
                continue

            # ケース2: 通常のワークシート上にある図形・グラフ・埋め込みオブジェクト。
            # 単純な貼り付け画像（msoPicture）はopenpyxl側で抽出済みなので除外し、
            # それ以外（オートシェイプ／グループ化された図形／SmartArt／ネイティブグラフ／
            # 埋め込みOLEオブジェクト等）は「構成図」の一部である可能性が高いのでまとめて画像化する。
            shape_count = sheet.Shapes.Count
            target_names = []
            for i in range(1, shape_count + 1):
                shp = sheet.Shapes.Item(i)
                if shp.Type != msoPicture:
                    target_names.append(shp.Name)

            if not target_names:
                continue

            try:
                # 複数の図形を一括選択してコピーすると、位置関係を保ったまま1枚の画像になる
                # （バラバラに描かれた矢印・箱・テキストで構成される構成図でもまとめて取れる）
                shp_range = sheet.Shapes.Range(target_names)
                shp_range.Copy()
                time.sleep(0.3)  # クリップボードへの反映待ち（自動化時のタイミング対策）

                temp_chart_obj = sheet.ChartObjects().Add(0, 0, 600, 400)
                temp_chart_obj.Chart.Paste()
                img_index += 1
                img_path = os.path.join(tmp_dir, f"diagram_{img_index}.png")
                temp_chart_obj.Chart.Export(img_path, "PNG")
                temp_chart_obj.Delete()

                with open(img_path, "rb") as img_f:
                    images.append({
                        "data": base64.b64encode(img_f.read()).decode("utf-8"),
                        "mime": "image/png",
                    })
            except Exception as shape_e:
                st.warning(f"「{filename}」の図形抽出中にエラーが発生しました（シート: {sheet.Name}）: {shape_e}")
    except Exception as e:
        st.warning(f"「{filename}」の図形・オブジェクト抽出中にエラーが発生しました: {e}")
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


# ★新機能: Windows＋Excelが無い環境（Streamlit Community CloudなどのLinux）向けの代替手段。
# LibreOffice（soffice）をヘッドレスモードで動かし、シートを「見た目どおり」にPDF化 →
# さらにページ単位のPNG画像に変換する。個々の図形を1つずつ判別するわけではなく、
# シート全体をレンダリングしてスクリーンショット的に画像化するので、
# オートシェイプ・SmartArt・グラフ・埋め込みオブジェクトなど種類を問わず「見えているもの」を丸ごと拾える。
# 前提: システムにLibreOfficeとpoppler-utils（pdftoppmコマンド）が入っていること。
# Streamlit Community Cloudの場合は、リポジトリ直下に置く packages.txt に
#   libreoffice
#   poppler-utils
# と書いておくと、デプロイ時に自動でaptインストールされる。
def extract_excel_pages_via_libreoffice(file_bytes, filename):
    images = []

    soffice_path = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_path:
        st.info(
            f"この環境（{platform.system()}）にはLibreOfficeが見つからないため、"
            "シート全体の画像化はスキップしました。"
            "Streamlit Community Cloudの場合は、リポジトリに `packages.txt` を追加し、"
            "`libreoffice` と `poppler-utils` を1行ずつ記載してから再デプロイしてください。"
        )
        return images

    tmp_dir = tempfile.mkdtemp(prefix="xlsx_lo_")
    tmp_xlsx_path = os.path.join(tmp_dir, filename)
    with open(tmp_xlsx_path, "wb") as f:
        f.write(file_bytes)

    try:
        # xlsx -> PDF（シートのレイアウト・図形・グラフを含めて見た目どおりに変換）
        # 複数プロセスが同時に動いても衝突しないよう、専用のユーザープロファイルを都度使う
        user_profile_dir = os.path.join(tmp_dir, "lo_profile")
        result = subprocess.run(
            [
                soffice_path,
                "--headless",
                "--norestore",
                f"-env:UserInstallation=file://{user_profile_dir}",
                "--convert-to", "pdf",
                "--outdir", tmp_dir,
                tmp_xlsx_path,
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )

        pdf_path = os.path.join(tmp_dir, os.path.splitext(os.path.basename(tmp_xlsx_path))[0] + ".pdf")
        if not os.path.exists(pdf_path):
            st.warning(
                f"「{filename}」のPDF変換に失敗しました: {result.stderr.strip() or result.stdout.strip()}"
            )
            return images

        pdftoppm_path = shutil.which("pdftoppm")
        if not pdftoppm_path:
            st.warning(
                "poppler-utils（pdftoppmコマンド）が見つからないため、PDFの画像化をスキップしました。"
                "packages.txt に `poppler-utils` を追加してください。"
            )
            return images

        page_prefix = os.path.join(tmp_dir, "page")
        subprocess.run(
            [pdftoppm_path, "-png", "-r", "150", pdf_path, page_prefix],
            capture_output=True,
            text=True,
            timeout=120,
        )

        for fname in sorted(os.listdir(tmp_dir)):
            if fname.startswith("page") and fname.endswith(".png"):
                with open(os.path.join(tmp_dir, fname), "rb") as img_f:
                    images.append({
                        "data": base64.b64encode(img_f.read()).decode("utf-8"),
                        "mime": "image/png",
                    })
    except Exception as e:
        st.warning(f"「{filename}」のLibreOffice変換中にエラーが発生しました: {e}")
    finally:
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

                # ★新機能: Excelの「グラフ」「図形（オートシェイプ／SmartArt／グループ化した図形）」
                # 「挿入した埋め込みオブジェクト」で作られた構成図を画像化する。
                # Windows＋Excelがある環境ではCOM経由（図形単位で高精度に切り出せる）、
                # それ以外（Streamlit Community CloudなどのLinux）ではLibreOffice経由（シート全体を画像化）
                # にフォールバックする。
                uploaded_file.seek(0)
                file_bytes = uploaded_file.read()
                if platform.system() == "Windows":
                    diagram_images = extract_excel_diagrams_via_com(file_bytes, uploaded_file.name)
                else:
                    diagram_images = extract_excel_pages_via_libreoffice(file_bytes, uploaded_file.name)
                if diagram_images:
                    st.session_state.extracted_images.extend(diagram_images)
                    st.success(
                        f"「{uploaded_file.name}」内の図形・構成図 {len(diagram_images)} 件をAIの「目」として抽出しました"
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

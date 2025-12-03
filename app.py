import streamlit as st
import os
import fitz  # PyMuPDF
import zipfile
import io
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- 設定 ---
SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'

def get_credentials():
    """認証情報を取得・更新する"""
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

def create_slides(uploaded_file, save_images):
    """メイン処理：PDFを受け取り、スライド作成と画像ZIP返却を行う"""
    creds = get_credentials()
    service_slides = build('slides', 'v1', credentials=creds)
    service_drive = build('drive', 'v3', credentials=creds)

    # アップロードされたファイルを一時保存
    pdf_path = "temp_uploaded.pdf"
    with open(pdf_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # PDFを開く
    doc = fitz.open(pdf_path)
    total_pages = len(doc)

    st.info(f"PDFを読み込みました。全 {total_pages} ページを処理します...")
    progress_bar = st.progress(0)

    # 画像保存用のZIP準備
    zip_buffer = io.BytesIO()

    # スライド作成
    presentation_title = os.path.splitext(uploaded_file.name)[0]
    presentation = service_slides.presentations().create(body={'title': presentation_title}).execute()
    presentation_id = presentation.get('presentationId')

    # デフォルトの1枚目ID取得
    first_slide_id = presentation.get('slides')[0].get('objectId')

    requests = []
    temp_images = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for i, page in enumerate(doc):
            # 画像化
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            image_filename = f"slide_{i+1:03}.png"
            pix.save(image_filename)
            temp_images.append(image_filename)

            # ZIPに追加（オプションの場合）
            if save_images:
                zip_file.write(image_filename)

            # Google Driveへアップロード
            file_metadata = {'name': image_filename}
            media = MediaFileUpload(image_filename, mimetype='image/png')
            drive_file = service_drive.files().create(body=file_metadata, media_body=media, fields='id').execute()
            file_id = drive_file.get('id')

            # 公開設定
            service_drive.permissions().create(
                fileId=file_id,
                body={'type': 'anyone', 'role': 'reader'}
            ).execute()

            image_url = f'https://drive.google.com/uc?id={file_id}&export=download'

            # スライド追加リクエスト
            page_id = f'slide_page_{i}'
            requests.append({
                'createSlide': {
                    'objectId': page_id,
                    'slideLayoutReference': {'predefinedLayout': 'BLANK'}
                }
            })
            requests.append({
                'createImage': {
                    'url': image_url,
                    'elementProperties': {
                        'pageObjectId': page_id,
                        'size': {'width': {'magnitude': 720, 'unit': 'PT'}, 'height': {'magnitude': 405, 'unit': 'PT'}},
                        'transform': {'scaleX': 1, 'scaleY': 1, 'translateX': 0, 'translateY': 0, 'unit': 'PT'}
                    }
                }
            })

            # プログレスバー更新
            progress_bar.progress((i + 1) / total_pages)

    # スライド実行
    requests.append({'deleteObject': {'objectId': first_slide_id}})
    if requests:
        service_slides.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()

    # 後始末
    doc.close()
    for f in temp_images:
        os.remove(f)
    os.remove(pdf_path)

    slide_url = f"https://docs.google.com/presentation/d/{presentation_id}"
    return slide_url, zip_buffer

# --- Streamlit UI ---
st.set_page_config(page_title="PDF to Google Slides", page_icon="📊")

st.title("📄 PDF to Google Slides Converter")
st.markdown("NotebookLMなどで作成したPDFを、画像化してGoogleスライドに変換します。")

# ファイルアップロード
uploaded_file = st.file_uploader("PDFファイルをここにドロップしてください", type="pdf")

# オプション
save_images_option = st.checkbox("スライドごとの画像(PNG)もダウンロードする", value=True)

if uploaded_file is not None:
    if st.button("🚀 スライド作成を開始", type="primary"):
        try:
            with st.spinner('変換中... コーヒーでも飲んでお待ちください ☕'):
                slide_url, zip_data = create_slides(uploaded_file, save_images_option)

            st.success("🎉 完了しました！")

            # スライドへのリンク
            st.markdown(f"### [🔗 Googleスライドを開く]({slide_url})")

            # 画像ZIPダウンロードボタン
            if save_images_option:
                st.download_button(
                    label="📥 画像(ZIP)をダウンロード",
                    data=zip_data.getvalue(),
                    file_name=f"{os.path.splitext(uploaded_file.name)[0]}_images.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
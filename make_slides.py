import os
import fitz  # PyMuPDF
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- 設定 ---
PDF_FILE = 'notebooklm_output.pdf' # 変換したいPDFファイル名
SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']

def get_credentials():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def pdf_to_slides_as_images():
    creds = get_credentials()
    service_slides = build('slides', 'v1', credentials=creds)
    service_drive = build('drive', 'v3', credentials=creds)

    # 1. 新規プレゼンテーション作成
    print("Googleスライドを作成中...")
    presentation = service_slides.presentations().create(body={'title': 'NotebookLM Slides (Image Ver)'}).execute()
    presentation_id = presentation.get('presentationId')

    # デフォルトの1枚目を削除するための準備
    first_slide_id = presentation.get('slides')[0].get('objectId')

    # 2. PDFを読み込んで画像化
    doc = fitz.open(PDF_FILE)
    print(f"PDFを開きました。全 {len(doc)} ページを処理します。")

    requests = []
    temp_files = []

    for i, page in enumerate(doc):
        # 高解像度で画像化 (matrix=2 で2倍の解像度。汚い場合は数値を上げる)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        image_filename = f"temp_slide_{i}.png"
        pix.save(image_filename)
        temp_files.append(image_filename)

        # Googleドライブにアップロード（APIがアクセスできるようにするため）
        print(f"ページ {i+1} をアップロード中...")
        file_metadata = {'name': image_filename}
        media = MediaFileUpload(image_filename, mimetype='image/png')
        drive_file = service_drive.files().create(body=file_metadata, media_body=media, fields='id').execute()
        file_id = drive_file.get('id')

        # 誰でも見れるように権限設定（一時的）
        # ※APIでスライドに貼るには画像が公開されているか、アクセストークンが必要ですが、
        # 簡略化のためDriveの画像を一時的にリンク共有可にします
        service_drive.permissions().create(
            fileId=file_id,
            body={'type': 'anyone', 'role': 'reader'}
        ).execute()

        image_url = f'https://drive.google.com/uc?id={file_id}&export=download'

        # スライド作成と画像配置のリクエスト
        page_id = f'slide_page_{i}'

        # 白紙スライドを追加
        requests.append({
            'createSlide': {
                'objectId': page_id,
                'slideLayoutReference': {'predefinedLayout': 'BLANK'}
            }
        })

        # 画像を画面いっぱいに配置
        # Googleスライドの標準サイズは 720pt x 405pt (16:9)
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

    # 最後にデフォルトの1枚目を削除
    requests.append({'deleteObject': {'objectId': first_slide_id}})

    # 3. まとめて実行
    print("スライドに画像を配置中...")
    if requests:
        service_slides.presentations().batchUpdate(presentationId=presentation_id, body={'requests': requests}).execute()

    # 4. 後始末（ローカルの画像削除）
    for f in temp_files:
        os.remove(f)

    print(f"完了しました！\nURL: https://docs.google.com/presentation/d/{presentation_id}")

if __name__ == '__main__':
    pdf_to_slides_as_images()
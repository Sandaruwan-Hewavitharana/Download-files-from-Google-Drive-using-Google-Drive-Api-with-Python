import flet as ft
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaIoBaseDownload
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from io import BytesIO
import os

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
TOKEN_FILE = 'token.json'
CREDENTIALS_FILE = 'credentials.json'

def create_service():
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
    
    return build('drive', 'v3', credentials=creds)

def main(page: ft.Page):
    page.title = "Google Drive Downloader"
    page.window_width = 400
    page.window_height = 600

    service = create_service()

    selected_files = []

    # Supported MIME types for export
    EXPORT_MIME_TYPES = {
        "application/vnd.google-apps.document": "application/pdf",
        "application/vnd.google-apps.spreadsheet": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.google-apps.presentation": "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    }

    def list_drive_files():
        files_list.controls.clear()  # Clear existing controls
        try:
            # Call the Drive v3 API to list files
            results = service.files().list(pageSize=10, fields="files(id, name, mimeType)").execute()
            items = results.get('files', [])

            if not items:
                files_list.controls.append(ft.Text("No files found."))
            else:
                for item in items:
                    checkbox = ft.Checkbox(label=item['name'], value=False, on_change=lambda e, item=item: select_file(e, item))
                    files_list.controls.append(checkbox)
            
            page.update()
        except Exception as e:
            files_list.controls.append(ft.Text(f"Error listing files: {e}"))
            page.update()

    def select_file(event, item):
        if event.control.value:
            selected_files.append(item)
        else:
            selected_files.remove(item)

    def download_selected_files():
        if not selected_files:
            return

        download_dialog.open = False
        page.update()

        for file in selected_files:
            download_file(file['id'], file['name'], file['mimeType'])

        selected_files.clear()

    def download_file(file_id, file_name, mime_type):
        try:
            # Check if file is a Google Workspace file and needs to be exported
            if mime_type in EXPORT_MIME_TYPES:
                request = service.files().export_media(fileId=file_id, mimeType=EXPORT_MIME_TYPES[mime_type])
                file_name += get_file_extension(EXPORT_MIME_TYPES[mime_type])
            else:
                request = service.files().get_media(fileId=file_id)

            fh = BytesIO()
            downloader = MediaIoBaseDownload(fd=fh, request=request)
            done = False

            download_progress_dialog.open = True
            download_progress_dialog.title = ft.Text(f"Downloading {file_name}")
            progress_bar.value = 0
            page.update()

            while not done:
                status, done = downloader.next_chunk()
                if status:
                    progress_bar.value = status.progress()
                    page.update()

            fh.seek(0)
            # Ensure binary mode is used for saving files
            with open(file_name, 'wb') as f:
                f.write(fh.read())
            
            download_progress_dialog.open = False
            page.update()
        except Exception as e:
            print(f"Error during download: {e}")

    def get_file_extension(mime_type):
        # Return appropriate file extension for MIME type
        return {
            "application/pdf": ".pdf",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx"
        }.get(mime_type, ".bin")

    def show_download_dialog():
        list_drive_files()
        download_dialog.open = True
        page.update()

    def show_upload_dialog():
        pass  # Placeholder for upload functionality

    download_dialog = ft.AlertDialog(
        title=ft.Text("Select Files to Download"),
        content=ft.Column(controls=[], scroll=True, expand=True),
        actions=[
            ft.TextButton("Download", on_click=lambda e: download_selected_files()),
            ft.TextButton("Close", on_click=lambda e: setattr(download_dialog, 'open', False))
        ],
        actions_alignment=ft.MainAxisAlignment.END,
        modal=True
    )

    download_progress_dialog = ft.AlertDialog(
        title=ft.Text("Download Progress"),
        content=ft.Column(controls=[
            ft.ProgressRing(value=0.0, color="green", width=100, height=100)
        ], scroll=True, expand=True),
        actions=[
            ft.TextButton("Close", on_click=lambda e: setattr(download_progress_dialog, 'open', False))
        ],
        actions_alignment=ft.MainAxisAlignment.END,
        modal=True
    )

    progress_bar = download_progress_dialog.content.controls[0]
    files_list = download_dialog.content
    download_dialog.content.controls.append(files_list)

    page.add(
        ft.Row([
            ft.TextButton("Download", on_click=lambda e: show_download_dialog(), expand=True),
            ft.TextButton("Upload", on_click=lambda e: show_upload_dialog(), expand=True)
        ], alignment=ft.MainAxisAlignment.CENTER, expand=True),
    )

    page.dialog = download_dialog

ft.app(target=main)

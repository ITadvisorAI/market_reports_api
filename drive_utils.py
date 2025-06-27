import os
import json
import traceback
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# === Google Drive Setup ===
drive_service = None
try:
    creds_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if creds_json:
        creds = service_account.Credentials.from_service_account_info(
            json.loads(creds_json),
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        drive_service = build("drive", "v3", credentials=creds)
    else:
        print("⚠️ GOOGLE_SERVICE_ACCOUNT_JSON not set; Drive uploads will fail.")
except Exception as e:
    print(f"❌ Failed to initialize Google Drive: {e}")
    traceback.print_exc()


def upload_to_drive(file_path: str, session_id: str, folder_id: str = None) -> str:
    """
    Uploads a file to a Google Drive folder.

    If folder_id is provided, the file is uploaded there.
    Otherwise, finds or creates a folder named after session_id.

    Returns the Drive "view" URL of the uploaded file,
    or None on failure.
    """
    try:
        if not drive_service:
            raise RuntimeError("Drive service not initialized")

        # Determine target folder
        if folder_id:
            target_folder = folder_id
        else:
            # Search for existing folder
            query = f"name='{session_id}' and mimeType='application/vnd.google-apps.folder'"
            resp = drive_service.files().list(q=query, fields="files(id)").execute()
            files = resp.get("files", [])
            if files:
                target_folder = files[0]["id"]
            else:
                folder_meta = {
                    "name": session_id,
                    "mimeType": "application/vnd.google-apps.folder"
                }
                created = drive_service.files().create(body=folder_meta, fields="id").execute()
                target_folder = created["id"]

        # Upload the file
        file_metadata = {"name": os.path.basename(file_path), "parents": [target_folder]}
        media = MediaFileUpload(file_path, resumable=True)
        uploaded = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()
        file_id = uploaded.get("id")
        return f"https://drive.google.com/file/d/{file_id}/view"

    except Exception as e:
        print(f"❌ Upload failed for {file_path}: {e}")
        traceback.print_exc()
        return None

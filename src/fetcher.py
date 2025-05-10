# fetcher.py
import os
import base64
import re
from pathlib import Path
from auth import authenticate_gmail

ATTACH_DIR = Path(__file__).resolve().parent.parent / "gradebooks"
EXT_REGEX = re.compile(r"\.xls[xm]?$", re.I)

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def list_unread_messages(service):
    query = "is:unread has:attachment from:*@ualberta.ca" # Adjust as needed
    # This query fetches unread emails with attachments from any @ualberta.ca address

    response = service.users().messages().list(userId='me', q=query).execute()
    return response.get('messages', [])

def download_attachments(service, message_id):
    message = service.users().messages().get(userId='me', id=message_id).execute()
    payload = message.get("payload", {})
    parts = payload.get("parts", [])

    for part in parts:
        filename = part.get("filename")
        if EXT_REGEX.search(filename or ""):
            body = part.get("body", {})
            att_id = body.get("attachmentId")

            if att_id:
                attachment = service.users().messages().attachments().get(
                    userId='me', messageId=message_id, id=att_id).execute()
                file_data = base64.urlsafe_b64decode(attachment['data'])

                # Save to disk
                ts = message["internalDate"]
                save_path = ATTACH_DIR / f"{ts}_{filename}"
                with open(save_path, 'wb') as f:
                    f.write(file_data)
                print(f"✔ Saved {filename} → {save_path}")
                # Mark as read (remove UNREAD label)
                service.users().messages().modify(
                    userId='me',
                    id=message_id,
                    body={'removeLabelIds': ['UNREAD']}
                ).execute()


def fetch_excel_attachments():
    ensure_dir(ATTACH_DIR)
    service = authenticate_gmail()
    messages = list_unread_messages(service)

    print(f"Found {len(messages)} unread email(s) with attachments.")
    for msg in messages:
        download_attachments(service, msg['id'])

if __name__ == "__main__":
    fetch_excel_attachments()

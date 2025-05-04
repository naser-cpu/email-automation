# src/outlook_fetcher.py
"""
Fetch unread Outlook messages and save any *.xls[x/m]* attachments
into the project‑level 'gradebooks/' directory.
"""

import os, re
from datetime import datetime
import win32com.client as win32

# ---------- config ----------
ATTACH_DIR   = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "gradebooks"))
INBOX_FOLDER = "Inbox"        # change if you want a sub‑folder (e.g., "Grades")
FILENAME_RE  = re.compile(r"\.xls[xm]?$", re.I)   # .xlsx, .xlsm, .xls
# -----------------------------

def ensure_dir(path: str):
    if not os.path.exists(path):
        os.makedirs(path)

def save_excel_attachments():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # -------- auto‑select store --------
    target_store = None
    for store in outlook.Folders:
        if "gmail" in store.Name.lower():      # tweak keyword if you ever switch accounts, can use any domain like 'ualberta'
            target_store = store
            break
    if target_store is None:
        target_store = outlook.GetDefaultFolder(6).Parent   # fallback = default store

    inbox = target_store.Folders["Inbox"]                   # Gmail inbox
    print(f"Scanning store: {target_store.Name!r} → Inbox")

 
    items = inbox.Items.Restrict("[Unread]=true")
    items.Sort("[ReceivedTime]", True)

    ensure_dir(ATTACH_DIR)
    saved = 0

    for item in items:
        # Check for attachments manually
        if item.Attachments.Count > 0:
            print(f"Processing email: {item.Subject}")
            for att in item.Attachments:
                if FILENAME_RE.search(att.FileName):
                    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
                    path = os.path.join(ATTACH_DIR, f"{ts}_{att.FileName}")
                    att.SaveAsFile(path)
                    print(f"   ✔  {att.FileName}  →  {path}")
                    saved += 1
            item.Unread = False          

    print(f"\nDone. {saved} file(s) saved from {target_store.Name!r}.")
    return saved


if __name__ == "__main__":
    save_excel_attachments()

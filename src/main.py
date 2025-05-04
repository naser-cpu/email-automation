# src/main.py
from fetcher import save_excel_attachments

def main():
    print(">>> STEP 1: fetching Excel attachments from Outlook …")
    save_excel_attachments()
    # later  chain: parse + load into SQLite

if __name__ == "__main__":
    main()

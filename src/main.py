# main.py
from pathlib import Path
from fetcher import fetch_excel_attachments
from parser import parse_gradebook
from database import *

GRADEBOOK_DIR = Path(__file__).resolve().parent.parent / "gradebooks"

def main():
    print("📥 Fetching unread gradebooks from @ualberta senders …")
    fetch_excel_attachments()

    print("Parsing gradebooks and saving to SQLite …")
    
    for xl_path in GRADEBOOK_DIR.glob("*.xls*"):
        with get_connection() as con:
            if was_file_processed(xl_path, con):
                print(f"Skipping already processed file: {xl_path.name}")
                continue
            try:
                parsed = parse_gradebook(xl_path)
                insert_course_and_grades(parsed)
                mark_file_processed(xl_path, con)
                con.commit()
            except Exception as e:
                print(f"⚠️  {xl_path.name}: {e}")

    print("✓ Done.")

if __name__ == "__main__":
    main()

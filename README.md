# email-automation
# 📬 Gmail Gradebook Automation

This Python project automates the process of extracting Excel gradebooks from Gmail, parsing student grades, and storing them in a structured SQLite database.

It simulates a real-world scenario where a department head or dean receives grade submissions via email and wants to automatically centralize and analyze them.

---

## 🚀 Features

- ✓ Connects securely to your Gmail account using OAuth
- ✓ Automatically fetches unread emails from `@ualberta.ca` with Excel attachments
- ✓ Parses course, instructor, and student data from a structured gradebook format
- ✓ Extracts student IDs and names from a separate sheet (`Names`) in the Excel file
- ✓ Computes and stores:
  - Course metadata
  - Per-student letter grades
  - File-level tracking to avoid duplicate imports
- ✓ Uses SQLite for lightweight local storage
- ✓ Idempotent: safe to re-run multiple times without duplicate data

---

## 📁 Project Structure
email-automation/
├── gradebooks/ # Saved Excel attachments
├── gradebooks.db # SQLite3 database
├── credentials.json # OAuth client secrets (from Google Cloud)
├── src/
│ ├── main.py # Main runner script
│ ├── auth.py # Gmail OAuth authentication
│ ├── fetcher.py # Downloads Excel attachments
│ ├── parser.py # Extracts data from Excel format
│ ├── database.py # SQLite schema + upserts + file tracking
| └──token.pickle # Gmail access token (auto-generated)
└── requirements.txt


---

## 🧠 Database Schema

### Courses

| id | name | instructor | mean_percentage | avg_letter_grade |
|----|------|------------|------------------|------------------|

### Students

| id | student_id | name | course_id | letter_grade |
|----|------------|------|-----------|---------------|

### ProcessedFiles

| id | filename | modified_at | sha256 |
|----|----------|--------------|--------|

- Ensures no file is processed twice (based on file name + timestamp + content hash)

---

## ⚙️ Setup Instructions

### 1. Clone and prepare environment

```bash
git clone https://github.com/yourusername/email-automation.git
cd email-automation
python3 -m venv .venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Enable Gmail API
- Go to Google Cloud Console

- Create a new project and enable the Gmail API

- Create OAuth 2.0 Credentials (Desktop App)

- Download credentials.json and place it in the root folder

- Naviagte to Google-Auth-Platform/Audience to add test user email

3. Run the pipeline
```bash
python3 src/main.py
```
On first run, you'll be prompted to log into your Gmail account via browser.

## Assumptions & Potential Improvements

### Assumptions

- The Excel files follow a **fixed structure**:
  - Course metadata is always in `B2`, `B3`
  - Student grades are always in `S11:S40`
  - Student names and IDs are in the `Names` sheet from `B10` and `C10` onward
- Email attachments are always **`.xlsx` or `.xlsm`**
- All messages come from **`@ualberta.ca`** addresses and are unread when processed
- Only **one course** is included per Excel file
- The Gmail inbox has sufficient quota and is not blocked by rate limiting
- The script is only run by **trusted users** (no hostile input expected)

---

### 🛡️ Security Considerations

- **SQLite and SQL Injection**
  - This project uses `sqlite3` with **parameterized queries (`?`)** to avoid SQL injection
  - Input from Excel files (e.g., student names) is never executed as SQL
  - No dynamic SQL construction or user inputs are passed directly

- **OAuth Token Storage**
  - Tokens are stored securely in `token.pickle` on disk
  - These tokens grant access to your Gmail inbox — do not share or commit them
  - Consider encrypting or rotating them if deploying long-term

- **File Handling**
  - The script does not currently sanitize file names (e.g., filenames with special characters)
  - No limit is placed on attachment size — large Excel files may impact memory

---

### 💡 Suggested Improvements

| Area                  | Enhancement                                                                 |
|-----------------------|------------------------------------------------------------------------------|
| **File deduplication**| Include full hash checking *and* original sender email in `ProcessedFiles`  |
| **Validation layer**  | Check for invalid grades, missing IDs, malformed letters                    |
| **Multi-course support** | Extend parser to handle multi-course Excel files                         |
| **Parallel processing**| Use threads or async I/O to handle large inboxes faster                    |
| **Deployment**        | Package into Docker or schedule using `cron` / `systemd`                    |

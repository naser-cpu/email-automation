# database.py
import sqlite3
from pathlib import Path
import hashlib
import os
from contextlib import contextmanager


DB_PATH = Path(__file__).resolve().parent.parent / "gradebooks.db"

DDL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS Courses (
    id INTEGER PRIMARY KEY,
    name TEXT,
    instructor TEXT,
    mean_percentage REAL,
    avg_letter_grade TEXT,
    UNIQUE(name, instructor)
);

CREATE TABLE IF NOT EXISTS Students (
    student_id TEXT PRIMARY KEY,
    name TEXT,
    course_id INTEGER,
    letter_grade TEXT,
    UNIQUE(student_id, course_id),
    FOREIGN KEY(course_id) REFERENCES Courses(id)
);

CREATE TABLE IF NOT EXISTS ProcessedFiles (
    id INTEGER PRIMARY KEY,
    filename TEXT UNIQUE,
    modified_at REAL,       -- os.path.getmtime()
    sha256 TEXT             -- file hash
);
"""

def get_file_hash(path: Path) -> str:
    hasher = hashlib.sha256()
    with open(path, 'rb') as f:
        while chunk := f.read(8192):
            hasher.update(chunk)
    return hasher.hexdigest()

def was_file_processed(path: Path, con: sqlite3.Connection) -> bool:
    mtime = os.path.getmtime(path)
    cur = con.cursor()
    cur.execute("""
        SELECT modified_at FROM ProcessedFiles
        WHERE filename = ?
    """, (path.name,))
    row = cur.fetchone()
    return row and abs(row[0] - mtime) < 1  # allow for ~1 sec clock diff

def mark_file_processed(path: Path, con: sqlite3.Connection):
    mtime = os.path.getmtime(path)
    sha = get_file_hash(path)
    cur = con.cursor()
    cur.execute("""
        INSERT INTO ProcessedFiles (filename, modified_at, sha256)
        VALUES (?, ?, ?)
        ON CONFLICT(filename) DO UPDATE SET
            modified_at = excluded.modified_at,
            sha256 = excluded.sha256
    """, (path.name, mtime, sha))

@contextmanager
def get_connection():
    first = not DB_PATH.exists()
    conn = sqlite3.connect(DB_PATH)
    try:
        if first:
            conn.executescript(DDL)
            conn.commit()
        yield conn
    finally:
        conn.close()

def insert_course_and_grades(parsed):
    with get_connection() as con:
        cur = con.cursor()

        cur.execute("""
            INSERT INTO Courses (name, instructor, mean_percentage, avg_letter_grade)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(name, instructor) DO UPDATE SET
                mean_percentage = excluded.mean_percentage,
                avg_letter_grade = excluded.avg_letter_grade
        """, (
            parsed["course_name"],
            parsed["instructor"],
            parsed["mean_percentage"],
            parsed["avg_letter_grade"]
        ))

        cur.execute("SELECT id FROM Courses WHERE name=? AND instructor=?",
                    (parsed["course_name"], parsed["instructor"]))
        course_id = cur.fetchone()[0]

        for _, row in parsed["grades_df"].iterrows():
            cur.execute("""
                INSERT INTO Students (student_id, name, course_id, letter_grade)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(student_id, course_id) DO UPDATE SET
                    name = excluded.name,
                    letter_grade = excluded.letter_grade
            """, (
                row.student_id,
                row.student_name,
                course_id,
                row.letter_grade
            ))

        con.commit()
    print(f"✓ Saved course '{parsed['course_name']}' ({len(parsed['grades_df'])} students)")

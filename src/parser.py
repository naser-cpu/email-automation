# parser.py
import openpyxl
from pathlib import Path
import pandas as pd

def parse_gradebook(path: Path):
    """
    Extracts:
      • course_name   (B2)
      • instructor    (B3)
      • mean_%        (R42)
      • avg_letter    (S42)
      • grades_df     (student_id, student_name, letter_grade)
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    # ------------- main (grade) sheet -------------
    main_ws = wb.active           # assumes first sheet has grades
    course_name  = main_ws["B2"].value
    instructor   = main_ws["B3"].value
    mean_percent = round(main_ws["R42"].value, 4)*100
    avg_letter   = main_ws["S42"].value

    # Pull letter grades (rows 11‑40, column S)
    letter_grades = []
    for r in range(11, 41):
        grade = main_ws[f"S{r}"].value
        letter_grades.append(grade)

    # ------------- Names sheet -------------
    names_ws = wb["Names"] if "Names" in wb.sheetnames else wb.worksheets[1]
    data = []
    row_idx = 0
    for r in range(10, 10 + len(letter_grades)):
        name  = names_ws[f"B{r}"].value
        sid   = names_ws[f"C{r}"].value
        grade = letter_grades[row_idx]

        # stop if the names column is blank
        if not name:
            break
        if grade is None:
            row_idx += 1
            continue

        data.append((sid, name, grade))
        row_idx += 1

    df = pd.DataFrame(data, columns=["student_id", "student_name", "letter_grade"])

    return {
        "course_name": course_name,
        "instructor": instructor,
        "mean_percentage": mean_percent,
        "avg_letter_grade": avg_letter,
        "grades_df": df
    }

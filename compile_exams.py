from pathlib import Path
import subprocess
import shutil

from openpyxl import load_workbook

# --------------------------------------------------
# Files and folders
# --------------------------------------------------

BASE_DIR = Path(__file__).parent

EXCEL_FILE = BASE_DIR / "ch5-banked-pts.xlsx"
MAIN_TEX = BASE_DIR / "main.tex"
STUDENT_DATA = BASE_DIR / "student-data.tex"
OUTPUT_DIR = BASE_DIR / "exams"

# --------------------------------------------------
# Map Excel columns to LaTeX variables
# --------------------------------------------------

score_mapping = {
    "Q1a": "QftcAMin",
    "Q1b": "QftcBMin",
    "Q1c": "QftcCMin",
    "Q1d": "QftcDMin",
    "Q2a": "QsubAMin",
    "Q2b": "QsubBMin",
    "Q2c": "QsubCMin",
    "Q2d": "QsubDMin",
}

# --------------------------------------------------
# Set up output directory
# --------------------------------------------------

OUTPUT_DIR.mkdir(exist_ok=True)

# --------------------------------------------------
# Read Excel workbook
# --------------------------------------------------

workbook = load_workbook(EXCEL_FILE, data_only=True)
sheet = workbook.active

rows = list(sheet.iter_rows(values_only=True))

headers = rows[0]

header_index = {
    header: index
    for index, header in enumerate(headers)
}

# --------------------------------------------------
# Compile one exam for each student
# --------------------------------------------------

for row in rows[2:11]: # make this 14 to render all exams 

    first_name = row[header_index["First Name"]]
    last_name = row[header_index["Last Name"]]

    print(f"Compiling exam for {first_name} {last_name}...")

    # ----------------------------------------------
    # Build student-data.tex
    # ----------------------------------------------

    latex_lines = [
        f"\\def\\studentfirstname{{{first_name}}}",
        f"\\def\\studentlastname{{{last_name}}}",
    ]

    for excel_column, latex_variable in score_mapping.items():

        score = row[header_index[excel_column]]

        # Treat an empty Excel cell as zero.
        if score is None:
            score = 0

        latex_lines.append(
            f"\\def\\{latex_variable}{{{score}}}"
        )

    STUDENT_DATA.write_text(
        "\n".join(latex_lines),
        encoding="utf-8"
    )

    # ----------------------------------------------
    # Compile LaTeX
    # ----------------------------------------------

    subprocess.run(
        [
            "pdflatex",
            "-interaction=nonstopmode",
            MAIN_TEX.name,
        ],
        cwd=BASE_DIR,
        check=True,
    )

    # ----------------------------------------------
    # Rename/copy resulting PDF
    # ----------------------------------------------

    generated_pdf = BASE_DIR / "main.pdf"

    safe_first = str(first_name).replace(" ", "_")
    safe_last = str(last_name).replace(" ", "_")

    output_pdf = (
        OUTPUT_DIR /
        f"Exam5_{safe_first}_{safe_last}.pdf"
    )

    shutil.copy2(generated_pdf, output_pdf)

print()
print("Finished compiling all exams.")
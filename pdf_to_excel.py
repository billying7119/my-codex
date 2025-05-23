import sys
from pathlib import Path

try:
    import PyPDF2
    import openpyxl
except ImportError as e:
    missing = e.name
    raise SystemExit(f"Missing required package: {missing}\nPlease install dependencies with 'pip install PyPDF2 openpyxl'.")


def pdf_to_excel(pdf_path: str, excel_path: str) -> None:
    """Extract text from *pdf_path* and save it line by line to *excel_path*."""
    pdf_reader = PyPDF2.PdfReader(open(pdf_path, "rb"))

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "PDF Text"

    row = 1
    for page in pdf_reader.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            sheet.cell(row=row, column=1, value=line)
            row += 1

    Path(excel_path).parent.mkdir(parents=True, exist_ok=True)
    workbook.save(excel_path)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python pdf_to_excel.py <input.pdf> <output.xlsx>")
        raise SystemExit(1)
    pdf_to_excel(sys.argv[1], sys.argv[2])

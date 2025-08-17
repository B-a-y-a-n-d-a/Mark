import re
import pytesseract
from PIL import Image
import openpyxl

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_answers_to_excel(image_path, excel_path="student_answers.xlsx"):
    img = Image.open(image_path)
    custom_config = r'--oem 1 --psm 6'
    extracted_text = pytesseract.image_to_string(img, config=custom_config)

    question_pattern = re.compile(
        r'(?:(?:Q\s?\d+)|(?:\d+\.)|(?:\d+\))|(?:\(\d+\))|(?:\([a-z]\)))',
        re.IGNORECASE
    )

    answers = {}
    current_q = None

    for line in extracted_text.splitlines():
        line = line.strip()
        if not line:
            continue
        match = question_pattern.match(line)
        if match:
            current_q = match.group().strip()
            answers[current_q] = line[len(match.group()):].strip()
        elif current_q:
            answers[current_q] += " " + line

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Answers"
    ws.append(["Question", "Student Answer"])

    for q, ans in answers.items():
        ws.append([q, ans.strip()])

    wb.save(excel_path)
    return excel_path

# Helper function for external use
def save_handwritten_answers(image_path, excel_path="student_answers.xlsx"):
    try:
        result_path = extract_answers_to_excel(image_path, excel_path)
        return result_path
    except Exception as e:
        raise RuntimeError(f"An error occurred: {e}")

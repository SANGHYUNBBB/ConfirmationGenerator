from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

EXCEL_PATH = BASE_DIR / "data" / "일임계약 리스트_통합.xlsx"
STAMP_IMAGE_PATH = BASE_DIR / "data" / "stamp.png"

OUTPUT_CUSTOMER_DIR = BASE_DIR / "pdf_customer"
OUTPUT_PT_DIR = BASE_DIR / "pdf_pt"

SHEET_NAME = "증액확인서"
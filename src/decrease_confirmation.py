import re
from pathlib import Path
from datetime import datetime

import pikepdf
import win32com.client as win32

from config import (
    EXCEL_PATH,
    OUTPUT_CUSTOMER_DIR,
    OUTPUT_PB_DIR,
    STAMP_IMAGE_PATH,
    LOGO_IMAGE_PATH,
)

SHEET_NAME = "감액 및 해지확인서"

# =========================
# 수동 조정용 상수
# =========================

# 로고 위치/크기 (페이지 기준)
LOGO_LEFT_CM = 14.9
LOGO_TOP_CM = 0.6
LOGO_WIDTH_CM = 4.8
LOGO_HEIGHT_CM = 1.4

STAMP_LEFT_CM = 14
STAMP_TOP_CM = 25
STAMP_SIZE_CM = 2.3

def clean_filename(value: str) -> str:
    value = str(value).strip()
    return re.sub(r'[\\/:*?"<>|]', "_", value)


def normalize_birth_password(value) -> str:
    text = str(value).strip()
    digits = re.sub(r"\D", "", text)
    return digits


def cm_to_points(cm: float) -> float:
    """
    Word 여백/페이지 크기 설정용 cm → point 변환
    """
    return cm / 2.54 * 72

def add_floating_image_by_page(
    document,
    image_path: Path,
    left_cm: float,
    top_cm: float,
    width_cm: float | None = None,
    height_cm: float | None = None,
):
    """
    페이지 기준 절대 위치에 이미지를 floating shape로 삽입합니다.
    텍스트 앞 배치라서 공간을 차지하지 않습니다.
    """

    if not image_path.exists():
        raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {image_path}")

    # Word 상수
    wdWrapFront = 3
    wdRelativeHorizontalPositionPage = 1
    wdRelativeVerticalPositionPage = 1

    width = -1 if width_cm is None else cm_to_points(width_cm)
    height = -1 if height_cm is None else cm_to_points(height_cm)

    shape = document.Shapes.AddPicture(
        FileName=str(image_path),
        LinkToFile=False,
        SaveWithDocument=True,
        Left=cm_to_points(left_cm),
        Top=cm_to_points(top_cm),
        Width=width,
        Height=height,
    )

    # 공간 차지하지 않도록 텍스트 앞
    shape.WrapFormat.Type = wdWrapFront

    # 핵심: 페이지 기준 좌표로 고정
    shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage

    shape.Left = cm_to_points(left_cm)
    shape.Top = cm_to_points(top_cm)

    return shape


def add_stamp_image_to_word(document):
    """
    도장을 페이지 기준 위치에 삽입합니다.
    공간을 차지하지 않고, 가로/세로 크기를 명시해서 깨짐을 방지합니다.
    """

    if not STAMP_IMAGE_PATH.exists():
        raise FileNotFoundError(f"도장 이미지 파일을 찾을 수 없습니다: {STAMP_IMAGE_PATH}")

    wdWrapFront = 3
    wdRelativeHorizontalPositionPage = 1
    wdRelativeVerticalPositionPage = 1

    shape = document.Shapes.AddPicture(
        FileName=str(STAMP_IMAGE_PATH),
        LinkToFile=False,
        SaveWithDocument=True,
        Left=cm_to_points(STAMP_LEFT_CM),
        Top=cm_to_points(STAMP_TOP_CM),
        Width=cm_to_points(STAMP_SIZE_CM),
        Height=cm_to_points(STAMP_SIZE_CM),
    )

    # 텍스트 앞: 공간 차지하지 않음
    shape.WrapFormat.Type = wdWrapFront

    # 페이지 기준 좌표로 고정
    shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage

    # 앞으로 가져오기
    shape.ZOrder(0)

    return shape
def add_logo_image_to_word(document):
    """
    우측 상단에 로고를 삽입합니다.
    공간을 차지하지 않도록 floating shape로 넣습니다.
    """

    shape = add_floating_image_by_page(
        document=document,
        image_path=LOGO_IMAGE_PATH,
        left_cm=LOGO_LEFT_CM,
        top_cm=LOGO_TOP_CM,
        width_cm=LOGO_WIDTH_CM,
        height_cm=LOGO_HEIGHT_CM,
    )

    return shape

def create_decrease_word_from_excel(
    account_no: str,
    valuation_amount: int | float,
    withdrawal_total: int | float,
    auto_transfer_yn: str,
    docx_path: Path
):
    """
    감액확인서 Word 생성

    감액 및 해지확인서 시트:
    K2  = 계좌번호
    P10 = 평가금액
    P11 = 인출금액합계
    P18 = 자동이체여부(Y/N)

    복사 범위:
    A38:G70

    파일명:
    N22
    """

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    workbook = None
    document = None

    try:
        workbook = excel.Workbooks.Open(str(EXCEL_PATH))
        sheet = workbook.Worksheets(SHEET_NAME)

        # 1. K2 계좌번호 입력
        sheet.Range("K2").Value = account_no

        # 1-1. L4:S4 복사해서 L8:S8에 값 붙여넣기
        sheet.Range("L4:S4").Copy()
        sheet.Range("L8:S8").PasteSpecial(Paste=-4163)  # xlPasteValues

        # 2. P10 평가금액 입력
        sheet.Range("P10").Value = valuation_amount

        # 3. P11 인출금액합계 입력
        sheet.Range("P11").Value = withdrawal_total

        # 4. P18 자동이체여부 입력
        auto_transfer_yn = auto_transfer_yn.strip().upper()

        if auto_transfer_yn not in ["Y", "N"]:
            raise ValueError("자동이체여부는 Y 또는 N만 입력할 수 있습니다.")

        sheet.Range("P18").Value = auto_transfer_yn

        # 5. 수식 재계산
        workbook.RefreshAll()
        excel.CalculateFullRebuild()

        # 6. 긴 공백 제거
        for cell in sheet.Range("A38:G70"):
            value = cell.Value

            if isinstance(value, str):
                value = re.sub(r"[ \u00A0]{2,}", " ", value)
                cell.Value = value

        # 7. 고객용 PDF 비밀번호: Q7 생년월일
        birth_value = sheet.Range("Q8").Value
        customer_password = normalize_birth_password(birth_value)

        if not customer_password:
            raise ValueError("Q8 셀의 생년월일 값이 비어 있어서 PDF 비밀번호를 만들 수 없습니다.")

        # 8. 파일명: N22 값 사용, yymmdd는 오늘 날짜로 교체
        filename_value = sheet.Range("N22").Value
        today_yymmdd = datetime.today().strftime("%y%m%d")

        base_filename = str(filename_value).replace("yymmdd", today_yymmdd)
        base_filename = clean_filename(base_filename)

        # 9. Word 새 문서 생성
        document = word.Documents.Add()

        # 10. 페이지 설정
        document.PageSetup.PageWidth = cm_to_points(21)
        document.PageSetup.PageHeight = cm_to_points(29.7)

        document.PageSetup.TopMargin = cm_to_points(1)
        document.PageSetup.BottomMargin = cm_to_points(0.5)
        document.PageSetup.LeftMargin = cm_to_points(1.2)
        document.PageSetup.RightMargin = cm_to_points(1.2)

        # 11. 엑셀 범위 복사
        sheet.Range("A38:G70").Copy()

        # 12. Word에 붙여넣기
        word.Selection.PasteExcelTable(
            False,  # LinkedToExcel
            False,  # WordFormatting
            False   # RTF
        )

        # 13. 표 레이아웃 > 자동 맞춤 > 창에 자동으로 맞춤
        # 13. 표 레이아웃 > 자동 맞춤 > 창에 자동으로 맞춤
        table = document.Tables(1)
        table.AutoFitBehavior(2)  # wdAutoFitWindow
        table.Rows.Alignment = 1  # wdAlignRowCenter

        table.Range.ParagraphFormat.SpaceBefore = 11
        table.Range.ParagraphFormat.SpaceAfter = 0
        table.Range.ParagraphFormat.LineSpacingRule = 0
        table.Rows.HeightRule = 0

        # 14. 로고 삽입 (우측 상단)
        add_logo_image_to_word(document)

        # 15. 도장 삽입
        add_stamp_image_to_word(document)

        # 16. docx 저장
        document.SaveAs2(
            str(docx_path),
            FileFormat=16
        )

        return customer_password, base_filename

    finally:
        if document is not None:
            document.Close(False)

        word.Quit()

        if workbook is not None:
            workbook.Close(SaveChanges=False)

        excel.Quit()


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path):
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    document = None

    try:
        document = word.Documents.Open(str(docx_path))
        document.ExportAsFixedFormat(
            OutputFileName=str(pdf_path),
            ExportFormat=17  # wdExportFormatPDF
        )

    finally:
        if document is not None:
            document.Close(False)

        word.Quit()


def encrypt_pdf(input_pdf_path: Path, output_pdf_path: Path, password: str):
    with pikepdf.open(input_pdf_path) as pdf:
        pdf.save(
            output_pdf_path,
            encryption=pikepdf.Encryption(
                owner=password,
                user=password,
                R=4
            )
        )


def generate_decrease_confirmation(
    account_no: str,
    valuation_amount: int | float,
    withdrawal_total: int | float,
    auto_transfer_yn: str
):
    """
    감액확인서 PDF 생성 메인 함수

    1. 고객용 PDF
       - 저장 위치: pdf_customer
       - 비밀번호: Q7 생년월일

    2. PB용 PDF
       - 저장 위치: pdf_pb
       - 비밀번호: 오늘 날짜 yymmdd
    """

    OUTPUT_CUSTOMER_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_PB_DIR.mkdir(parents=True, exist_ok=True)

    temp_docx_path = OUTPUT_CUSTOMER_DIR / "temp_decrease_confirmation.docx"

    customer_password, base_filename = create_decrease_word_from_excel(
        account_no=account_no,
        valuation_amount=valuation_amount,
        withdrawal_total=withdrawal_total,
        auto_transfer_yn=auto_transfer_yn,
        docx_path=temp_docx_path
    )

    pb_password = datetime.today().strftime("%y%m%d")

    temp_pdf_path = OUTPUT_CUSTOMER_DIR / f"{base_filename}_temp.pdf"

    customer_pdf_path = OUTPUT_CUSTOMER_DIR / f"{base_filename}.pdf"
    pb_pdf_path = OUTPUT_PB_DIR / f"{base_filename}.pdf"

    # 1. Word → 임시 PDF 변환
    convert_docx_to_pdf(temp_docx_path, temp_pdf_path)

    # 2. 고객용 PDF: Q7 생년월일 비밀번호
    encrypt_pdf(
        input_pdf_path=temp_pdf_path,
        output_pdf_path=customer_pdf_path,
        password=customer_password
    )

    # 3. PB용 PDF: 오늘 날짜 yymmdd 비밀번호
    encrypt_pdf(
        input_pdf_path=temp_pdf_path,
        output_pdf_path=pb_pdf_path,
        password=pb_password
    )

    # 4. 임시 파일 삭제
    temp_pdf_path.unlink(missing_ok=True)

    # Word 파일 확인하고 싶으면 아래 줄 주석 처리
    temp_docx_path.unlink(missing_ok=True)

    print(f"고객용 PDF 생성 완료: {customer_pdf_path}")
    print(f"고객용 PDF 비밀번호: {customer_password}")

    print(f"PB용 PDF 생성 완료: {pb_pdf_path}")
    print(f"PB용 PDF 비밀번호: {pb_password}")


if __name__ == "__main__":
    account_no = input("계좌번호를 입력하세요: ").strip()

    valuation_amount = input("평가금액을 입력하세요: ").strip()
    valuation_amount = int(valuation_amount.replace(",", ""))

    withdrawal_total = input("인출금액합계를 입력하세요: ").strip()
    withdrawal_total = int(withdrawal_total.replace(",", ""))

    auto_transfer_yn = input("자동이체여부를 입력하세요(Y/N): ").strip().upper()

    generate_decrease_confirmation(
        account_no=account_no,
        valuation_amount=valuation_amount,
        withdrawal_total=withdrawal_total,
        auto_transfer_yn=auto_transfer_yn
    )
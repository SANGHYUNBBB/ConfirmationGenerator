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
    ACCOUNT_IMAGE_PATH,
)

SHEET_NAME = "연장확인서"


# =========================
# 수동 조정용 상수
# =========================

# 로고 위치/크기
LOGO_LEFT_CM = 16.3
LOGO_TOP_CM = 0.5
LOGO_WIDTH_CM = 3.8
LOGO_HEIGHT_CM = 1.2

# 도장 위치/크기 - 회사명 문구 기준 상대 위치
STAMP_OFFSET_X_CM = 13.8
STAMP_OFFSET_Y_CM = -1.9
STAMP_SIZE_CM = 2


def clean_filename(value: str) -> str:
    value = str(value).strip()
    return re.sub(r'[\\/:*?"<>|]', "_", value)


def normalize_birth_password(value) -> str:
    """
    생년월일 값을 PDF 비밀번호용 YYMMDD 문자열로 변환합니다.

    예:
    - 860104.0     -> 860104
    - 860104       -> 860104
    - 010104       -> 010104
    - 19860104     -> 860104
    - 1986-01-04   -> 860104
    - 1986.01.04   -> 860104
    - datetime     -> yymmdd
    """

    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.strftime("%y%m%d")

    if isinstance(value, (int, float)):
        if float(value).is_integer():
            text = str(int(value))
        else:
            text = str(value)
    else:
        text = str(value).strip()

    digits = re.sub(r"\D", "", text)

    if len(digits) == 8:
        return digits[2:]

    if 1 <= len(digits) < 6:
        return digits.zfill(6)

    if len(digits) == 7 and digits.endswith("0"):
        return digits[:-1]

    return digits


def cm_to_points(cm: float) -> float:
    """
    Word 여백/페이지 크기 설정용 cm → point 변환
    """
    return cm / 2.54 * 72


def find_filled_source_row(sheet, start_row: int = 2, end_row: int = 5) -> int:
    """
    K2에 계좌번호를 입력한 뒤,
    L2:S5 범위 중 실제로 값이 작성된 행을 찾습니다.

    상품마다 결과가 L2:S2, L3:S3, L4:S4, L5:S5 중
    다른 행에 나타날 수 있으므로, 가장 많이 채워진 행을 선택합니다.
    """

    best_row = None
    best_count = 0

    for row in range(start_row, end_row + 1):
        filled_count = 0

        for col in range(12, 20):  # L=12, S=19
            value = sheet.Cells(row, col).Value

            if value is not None and str(value).strip() != "":
                filled_count += 1

        if filled_count > best_count:
            best_count = filled_count
            best_row = row

    if best_row is None or best_count == 0:
        raise ValueError("L2:S5 범위에서 작성된 행을 찾지 못했습니다.")

    return best_row


def find_text_range(document, target_text: str):
    wdFindContinue = 1

    rng = document.Content
    finder = rng.Find
    finder.ClearFormatting()
    finder.Text = target_text
    finder.Forward = True
    finder.Wrap = wdFindContinue

    found = finder.Execute()

    if not found:
        raise ValueError(f"문서에서 '{target_text}' 문구를 찾지 못했습니다.")

    return rng.Duplicate


def add_floating_image_fixed(
    document,
    image_path: Path,
    anchor_range,
    left_cm: float,
    top_cm: float,
    width_cm: float,
    height_cm: float,
):
    """
    anchor는 특정 문단/문구에 잡되,
    위치는 페이지 기준 절대 좌표로 고정합니다.
    표 셀 영향은 LayoutInCell=False로 방지합니다.
    """

    if not image_path.exists():
        raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {image_path}")

    wdWrapFront = 3
    wdRelativeHorizontalPositionPage = 1
    wdRelativeVerticalPositionPage = 1

    shape = document.Shapes.AddPicture(
        FileName=str(image_path),
        LinkToFile=False,
        SaveWithDocument=True,
        Anchor=anchor_range,
    )

    shape.WrapFormat.Type = wdWrapFront
    shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shape.RelativeVerticalPosition = wdRelativeVerticalPositionPage

    try:
        shape.LayoutInCell = False
    except Exception:
        pass

    shape.LockAnchor = True

    shape.Left = cm_to_points(left_cm)
    shape.Top = cm_to_points(top_cm)
    shape.Width = cm_to_points(width_cm)
    shape.Height = cm_to_points(height_cm)

    shape.ZOrder(0)

    return shape


def add_logo_image_to_word(document):
    """
    우측 상단에 로고를 삽입합니다.
    """
    anchor_range = document.Paragraphs(1).Range

    return add_floating_image_fixed(
        document=document,
        image_path=LOGO_IMAGE_PATH,
        anchor_range=anchor_range,
        left_cm=LOGO_LEFT_CM,
        top_cm=LOGO_TOP_CM,
        width_cm=LOGO_WIDTH_CM,
        height_cm=LOGO_HEIGHT_CM,
    )

def add_account_image_page(document):
    """
    Word 문서 마지막에 새 페이지를 추가하고,
    data/account/account.png 이미지를 첨부합니다.
    """

    if not ACCOUNT_IMAGE_PATH.exists():
        raise FileNotFoundError(f"계좌 이미지 파일을 찾을 수 없습니다: {ACCOUNT_IMAGE_PATH}")

    # Word 상수
    wdStory = 6
    wdPageBreak = 7
    wdAlignParagraphCenter = 1

    app = document.Application

    # 문서 끝으로 이동
    app.Selection.EndKey(Unit=wdStory)

    # 새 페이지 추가
    app.Selection.InsertBreak(Type=wdPageBreak)

    # 가운데 정렬
    app.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

    # 이미지 삽입
    inline_shape = app.Selection.InlineShapes.AddPicture(
        FileName=str(ACCOUNT_IMAGE_PATH),
        LinkToFile=False,
        SaveWithDocument=True,
    )

    # 이미지 크기 조정
    inline_shape.LockAspectRatio = True
    inline_shape.Width = cm_to_points(16.0)

    return inline_shape


def add_stamp_image_to_word(document):
    """
    '주식회사 플레인바닐라투자자문' 문구를 기준으로
    도장을 상대 위치에 삽입합니다.

    문서 내용 길이가 달라져도 도장이 회사명 문구 주변에 따라붙습니다.
    """

    if not STAMP_IMAGE_PATH.exists():
        raise FileNotFoundError(f"도장 이미지 파일을 찾을 수 없습니다: {STAMP_IMAGE_PATH}")

    company_text = "주식회사 플레인바닐라투자자문"

    wdFindContinue = 1
    wdCollapseEnd = 0
    wdWrapFront = 3

    wdRelativeHorizontalPositionCharacter = 3
    wdRelativeVerticalPositionLine = 3

    company_range = document.Content
    find = company_range.Find
    find.ClearFormatting()
    find.Text = company_text
    find.Forward = True
    find.Wrap = wdFindContinue

    found = find.Execute()

    if not found:
        raise ValueError(f"문서에서 '{company_text}' 문구를 찾지 못했습니다.")

    anchor_range = company_range.Duplicate
    anchor_range.Collapse(wdCollapseEnd)

    shape = document.Shapes.AddPicture(
        FileName=str(STAMP_IMAGE_PATH),
        LinkToFile=False,
        SaveWithDocument=True,
        Anchor=anchor_range,
    )

    shape.WrapFormat.Type = wdWrapFront

    # 회사명 문구 기준 상대 위치
    shape.RelativeHorizontalPosition = wdRelativeHorizontalPositionCharacter
    shape.RelativeVerticalPosition = wdRelativeVerticalPositionLine

    shape.Left = cm_to_points(STAMP_OFFSET_X_CM)
    shape.Top = cm_to_points(STAMP_OFFSET_Y_CM)

    shape.Width = cm_to_points(STAMP_SIZE_CM)
    shape.Height = cm_to_points(STAMP_SIZE_CM)

    try:
        shape.LayoutInCell = False
    except Exception:
        pass

    shape.LockAnchor = True
    shape.ZOrder(0)

    return shape

def add_blank_line_before_date_text(document):
    """
    Word 문서 안에서 오늘 날짜 문구를 찾아
    그 앞에 빈 줄을 추가합니다.

    고객정보와 날짜 사이 간격을 만들기 위한 함수입니다.
    """

    date_text = datetime.today().strftime("%Y년 %m월 %d일")

    wdFindContinue = 1
    wdCollapseStart = 1

    find_range = document.Content
    find = find_range.Find
    find.ClearFormatting()
    find.Text = date_text
    find.Forward = True
    find.Wrap = wdFindContinue

    found = find.Execute()

    if not found:
        print(f"날짜 문구를 찾지 못했습니다: {date_text}")
        return False

    # 날짜 문구 시작 위치로 이동
    insert_range = find_range.Duplicate
    insert_range.Collapse(wdCollapseStart)

    # 날짜 앞에 빈 줄 추가
    insert_range.InsertBefore("\r\r\r")

    # print(f"날짜 앞 빈 줄 추가 완료: {date_text}")
    return True

def create_extension_word_from_excel(
    account_no: str,
    valuation_amount: int | float,
    auto_transfer_yn: str,
    docx_path: Path,
):
    """
    계약 연장확인서 Word 생성

    입력:
    K2  = 계좌번호
    L2:S5 중 작성된 행 → L8:S8
    P10 = 평가금액
    P14 = 자동이체여부(Y/N)

    복사 범위:
    A3:G31

    파일명:
    N19
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

        # K2 입력 후 수식 계산
        excel.CalculateFullRebuild()

        # 2. L2:S5 중 실제 값이 작성된 행을 찾아 L8:S8에 값 붙여넣기
        source_row = find_filled_source_row(sheet)

        sheet.Range(f"L{source_row}:S{source_row}").Copy()
        sheet.Range("L8:S8").PasteSpecial(Paste=-4163)  # xlPasteValues

        # 3. P10 평가금액 입력
        sheet.Range("P10").Value = valuation_amount

        # 4. P14 자동이체여부 입력
        auto_transfer_yn = auto_transfer_yn.strip().upper()

        if auto_transfer_yn not in ["Y", "N"]:
            raise ValueError("자동이체여부는 Y 또는 N만 입력할 수 있습니다.")

        sheet.Range("P14").Value = auto_transfer_yn

        # 5. 수식 재계산
        workbook.RefreshAll()
        excel.CalculateFullRebuild()

        # 6. 긴 공백 제거
        for cell in sheet.Range("A3:G31"):
            value = cell.Value

            if isinstance(value, str):
                value = re.sub(r"[ \u00A0]{2,}", " ", value)
                cell.Value = value

        # 7. 고객용 PDF 비밀번호: Q8 생년월일
        birth_value = sheet.Range("Q8").Value
        customer_password = normalize_birth_password(birth_value)

        if not customer_password:
            raise ValueError("Q8 셀의 생년월일 값이 비어 있어서 PDF 비밀번호를 만들 수 없습니다.")

        # 고객 이메일: R8
        email_value = sheet.Range("R8").Value
        customer_email = "" if email_value is None else str(email_value).strip()
       
        # 8. 파일명: N19 값 사용, yymmdd는 오늘 날짜로 교체
        filename_value = sheet.Range("N19").Value
        today_yymmdd = datetime.today().strftime("%y%m%d")

        base_filename = str(filename_value).replace("yymmdd", today_yymmdd)
        base_filename = clean_filename(base_filename)

        # 9. Word 새 문서 생성
        document = word.Documents.Add()

        # 10. 페이지 설정
        document.PageSetup.PageWidth = cm_to_points(21)
        document.PageSetup.PageHeight = cm_to_points(29.7)

        document.PageSetup.TopMargin = cm_to_points(-3)
        document.PageSetup.BottomMargin = cm_to_points(0.5)
        document.PageSetup.LeftMargin = cm_to_points(1.2)
        document.PageSetup.RightMargin = cm_to_points(1.2)

        # 11. 엑셀 범위 복사
        sheet.Range("A3:G31").Copy()

        # 12. Word에 붙여넣기
        word.Selection.PasteExcelTable(
            False,  # LinkedToExcel
            False,  # WordFormatting
            False,  # RTF
        )

        # 13. 표 레이아웃 > 자동 맞춤 > 창에 자동으로 맞춤
        table = document.Tables(1)
        table.AutoFitBehavior(2)
        table.Rows.Alignment = 1

        # 글자 크기는 건드리지 않음
        table.Range.ParagraphFormat.SpaceBefore = 5
        table.Range.ParagraphFormat.SpaceAfter = 5
        table.Range.ParagraphFormat.LineSpacingRule = 0
        table.Rows.HeightRule = 0

        # 고객정보와 날짜 사이 간격 추가
        add_blank_line_before_date_text(document)

        # 14. 로고 삽입
        add_logo_image_to_word(document)

        # 15. 도장 삽입
        add_stamp_image_to_word(document)

        # 16. 다음 페이지에 계좌 이미지 첨부
        add_account_image_page(document)

        # 17. docx 저장
        document.SaveAs2(
            str(docx_path),
            FileFormat=16,
        )

        return customer_password, base_filename, customer_email

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
            ExportFormat=17,
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
                R=4,
            ),
        )


def generate_extension_confirmation(
    account_no: str,
    valuation_amount: int | float,
    auto_transfer_yn: str,
):
    """
    계약 연장확인서 PDF 생성 메인 함수

    고객용 PDF:
    - 저장 위치: pdf_customer
    - 비밀번호: Q8 생년월일

    PB용 PDF:
    - 저장 위치: pdf_pb
    - 비밀번호: 오늘 날짜 yymmdd
    """

    OUTPUT_CUSTOMER_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_PB_DIR.mkdir(parents=True, exist_ok=True)

    temp_docx_path = OUTPUT_CUSTOMER_DIR / "temp_extension_confirmation.docx"

    customer_password, base_filename, customer_email = create_extension_word_from_excel(
        account_no=account_no,
        valuation_amount=valuation_amount,
        auto_transfer_yn=auto_transfer_yn,
        docx_path=temp_docx_path,
    )

    pb_password = datetime.today().strftime("%y%m%d")

    temp_pdf_path = OUTPUT_CUSTOMER_DIR / f"{base_filename}_temp.pdf"

    customer_pdf_path = OUTPUT_CUSTOMER_DIR / f"{base_filename}.pdf"
    pb_pdf_path = OUTPUT_PB_DIR / f"{base_filename}.pdf"

    # 1. Word → 임시 PDF 변환
    convert_docx_to_pdf(temp_docx_path, temp_pdf_path)

    # 2. 고객용 PDF
    encrypt_pdf(
        input_pdf_path=temp_pdf_path,
        output_pdf_path=customer_pdf_path,
        password=customer_password,
    )

    # 3. PB용 PDF
    encrypt_pdf(
        input_pdf_path=temp_pdf_path,
        output_pdf_path=pb_pdf_path,
        password=pb_password,
    )

    # 4. 임시 파일 삭제
    temp_pdf_path.unlink(missing_ok=True)

    # Word 파일 확인하고 싶으면 아래 줄 주석 처리
    temp_docx_path.unlink(missing_ok=True)

    print(f"고객용 PDF 생성 완료: {customer_pdf_path}")
    print(f"고객용 PDF 비밀번호: {customer_password}")
    print(f"고객 이메일: {customer_email}")
    print(f"PB용 PDF 생성 완료: {pb_pdf_path}")
    print(f"PB용 PDF 비밀번호: {pb_password}")


if __name__ == "__main__":
    account_no = input("계좌번호를 입력하세요: ").strip()

    valuation_amount = input("평가금액을 입력하세요: ").strip()
    valuation_amount = int(valuation_amount.replace(",", ""))

    auto_transfer_yn = input("자동이체여부를 입력하세요(Y/N): ").strip().upper()

    generate_extension_confirmation(
        account_no=account_no,
        valuation_amount=valuation_amount,
        auto_transfer_yn=auto_transfer_yn,
    )
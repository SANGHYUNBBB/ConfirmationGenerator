import re
from pathlib import Path
from datetime import datetime

import pikepdf
import win32com.client as win32

from config import EXCEL_PATH, OUTPUT_CUSTOMER_DIR, OUTPUT_PB_DIR, INCREASE_SHEET_NAME, STAMP_IMAGE_PATH


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
    1 inch = 2.54 cm, 1 inch = 72 points
    """
    return cm / 2.54 * 72

def create_word_from_excel(account_no: str, increase_amount: int | float, docx_path: Path):
    """
    엑셀 A4:G33 범위를 Word에 복붙한 뒤
    표 레이아웃 > 자동 맞춤 > 창에 자동으로 맞춤 적용
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
        sheet = workbook.Worksheets(INCREASE_SHEET_NAME)

        # 1. K2 계좌번호 입력
        sheet.Range("K2").Value = account_no

        # 2. L4:R4 복사해서 L7:R7에 값 붙여넣기
        sheet.Range("L4:R4").Copy()
        sheet.Range("L7:R7").PasteSpecial(Paste=-4163)  # xlPasteValues

        # 3. P9 증액금액 입력
        sheet.Range("P9").Value = increase_amount

        # 4. 수식 재계산
        workbook.RefreshAll()
        excel.CalculateFullRebuild()

        # 4-1. 안내 문구 안의 불필요한 긴 공백 제거
        for cell in sheet.Range("A4:G33"):
            value = cell.Value

            if isinstance(value, str) and "일임재산금액" in value:
                value = re.sub(r"[ \u00A0]{2,}", " ", value)
                cell.Value = value

        # 5. 비밀번호: Q7 생년월일
        birth_value = sheet.Range("Q7").Value
        password = normalize_birth_password(birth_value)

        if not password:
            raise ValueError("Q7 셀의 생년월일 값이 비어 있어서 PDF 비밀번호를 만들 수 없습니다.")

        # 6. 파일명: M12 값에서 yymmdd만 오늘 날짜로 교체
        filename_value = sheet.Range("M12").Value
        today_yymmdd = datetime.today().strftime("%y%m%d")

        base_filename = str(filename_value).replace("yymmdd", today_yymmdd)
        base_filename = clean_filename(base_filename)

        # 7. Word 새 문서 생성
        document = word.Documents.Add()

        # 8. 페이지 설정
        document.PageSetup.PageWidth = cm_to_points(21)
        document.PageSetup.PageHeight = cm_to_points(29.7)

        document.PageSetup.TopMargin = cm_to_points(1.5)
        document.PageSetup.BottomMargin = cm_to_points(1.5)
        document.PageSetup.LeftMargin = cm_to_points(1.5)
        document.PageSetup.RightMargin = cm_to_points(1.5)

        # 9. 엑셀 범위 복사
        sheet.Range("A4:G33").Copy()

        # 10. Word에 붙여넣기
        word.Selection.PasteExcelTable(
            False,  # LinkedToExcel
            False,  # WordFormatting
            False   # RTF
        )

        # 11. 표 레이아웃 > 자동 맞춤 > 창에 자동으로 맞춤
        table = document.Tables(1)
        table.AutoFitBehavior(2)  # wdAutoFitWindow

        # 표 가운데 정렬
        table.Rows.Alignment = 1  # wdAlignRowCenter


        # 11-1. 회사명 오른쪽에 도장 이미지 삽입
        add_stamp_image_to_word(document)
        # 12. docx 저장
        document.SaveAs2(
            str(docx_path),
            FileFormat=16  # wdFormatXMLDocument = docx
        )

        return password, base_filename

    finally:
        if document is not None:
            document.Close(False)

        word.Quit()

        if workbook is not None:
            workbook.Close(SaveChanges=False)

        excel.Quit()

def add_stamp_image_to_word(document):
    """
    회사명 오른쪽에 도장을 삽입합니다.
    1) 먼저 인라인으로 삽입해서 이미지 삽입을 확실히 성공시킴
    2) 바로 floating shape로 변환해서 공간 차지를 없앰
    3) 위치는 회사명 문구 근처 기준으로 조정
    """

    if not STAMP_IMAGE_PATH.exists():
        raise FileNotFoundError(f"도장 이미지 파일을 찾을 수 없습니다: {STAMP_IMAGE_PATH}")

    company_text = "주식회사 플레인바닐라투자자문"

    # Word 상수
    wdFindContinue = 1
    wdCollapseEnd = 0
    wdWrapFront = 3  # 텍스트 앞: 공간 차지 X, behind처럼 숨지 않음

    # 회사명 찾기
    find_range = document.Content
    find = find_range.Find
    find.ClearFormatting()
    find.Text = company_text
    find.Forward = True
    find.Wrap = wdFindContinue

    found = find.Execute()

    if not found:
        raise ValueError(f"Word 문서에서 '{company_text}' 문구를 찾지 못했습니다.")

    # 회사명 끝 위치로 이동
    insert_range = find_range.Duplicate
    insert_range.Collapse(wdCollapseEnd)

    # 일단 인라인으로 삽입
    inline_shape = document.InlineShapes.AddPicture(
        FileName=str(STAMP_IMAGE_PATH),
        LinkToFile=False,
        SaveWithDocument=True,
        Range=insert_range
    )

    # 도장 크기
    inline_shape.LockAspectRatio = True
    inline_shape.Height = cm_to_points(2.5)

    # 인라인 이미지를 floating shape로 변환
    shape = inline_shape.ConvertToShape()

    # 텍스트 앞 배치: 공간 차지 안 함
    shape.WrapFormat.Type = wdWrapFront

    # 핵심:
    # 페이지 절대좌표를 쓰지 말고, 방금 삽입된 위치 기준으로 조금만 이동
    shape.Left = cm_to_points(13.3)
    shape.Top = cm_to_points(-0.9)

    return shape

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


def generate_increase_confirmation(account_no: str, increase_amount: int | float):
    """
    증액확인서 PDF 생성 메인 함수

    1. 고객용 PDF
       - 저장 위치: pdf_customer
       - 비밀번호: Q7 생년월일

    2. PB용 PDF
       - 저장 위치: pdf_pb
       - 비밀번호: 오늘 날짜 yymmdd
    """

    OUTPUT_CUSTOMER_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_PB_DIR.mkdir(parents=True, exist_ok=True)

    temp_docx_path = OUTPUT_CUSTOMER_DIR / "temp_increase_confirmation.docx"

    customer_password, base_filename = create_word_from_excel(
        account_no=account_no,
        increase_amount=increase_amount,
        docx_path=temp_docx_path
    )

    # PB용 비밀번호: 오늘 날짜 yymmdd
    pb_password = datetime.today().strftime("%y%m%d")

    temp_pdf_path = OUTPUT_CUSTOMER_DIR / f"{base_filename}_temp.pdf"

    customer_pdf_path = OUTPUT_CUSTOMER_DIR / f"{base_filename}.pdf"
    pb_pdf_path = OUTPUT_PB_DIR / f"{base_filename}.pdf"

    # 1. Word → 임시 PDF 변환
    convert_docx_to_pdf(temp_docx_path, temp_pdf_path)

    # 2. 고객용 PDF 암호화
    encrypt_pdf(
        input_pdf_path=temp_pdf_path,
        output_pdf_path=customer_pdf_path,
        password=customer_password
    )

    # 3. PB용 PDF 암호화
    encrypt_pdf(
        input_pdf_path=temp_pdf_path,
        output_pdf_path=pb_pdf_path,
        password=pb_password
    )

    # 4. 임시 파일 삭제
    temp_pdf_path.unlink(missing_ok=True)
    temp_docx_path.unlink(missing_ok=True)

    print(f"고객용 PDF 생성 완료: {customer_pdf_path}")
    print(f"고객용 PDF 비밀번호: {customer_password}")

    print(f"PB용 PDF 생성 완료: {pb_pdf_path}")
    print(f"PB용 PDF 비밀번호: {pb_password}")


if __name__ == "__main__":
    account_no = input("계좌번호를 입력하세요: ").strip()
    increase_amount = input("증액금액을 입력하세요: ").strip()

    increase_amount = int(increase_amount.replace(",", ""))

    generate_increase_confirmation(
        account_no=account_no,
        increase_amount=increase_amount
    )    
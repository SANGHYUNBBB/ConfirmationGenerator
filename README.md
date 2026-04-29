# ConfirmationGenerator

엑셀에 정리된 고객/계약 정보를 기준으로 확인서 PDF를 자동 생성하는 프로그램입니다.

고객용 PDF와 PB용 PDF를 각각 생성하며, PDF 비밀번호도 자동으로 설정됩니다.

## 지원 문서

- 투자일임재산 변경확인서(증액)
- 투자일임재산 변경확인서(감액)
- 일임운용 정산 및 연장 내역서
- 일임 해지 및 정산 확인서

---


## 1. 프로젝트 다운로드

GitHub 저장소에서 프로젝트를 다운로드합니다.

### 방법 1. Git으로 다운로드

PowerShell 또는 명령 프롬프트에서 아래 명령어를 실행합니다.

```bash
git clone https://github.com/SANGHYUNBBB/ConfirmationGenerator.git
```

다운로드 후 프로젝트 폴더로 이동합니다.

```bash
cd ConfirmationGenerator
```

### 방법 2. ZIP 파일로 다운로드

Git 사용이 어렵다면 아래 방법으로 다운로드합니다.

1. GitHub 저장소 접속
2. 초록색 `Code` 버튼 클릭
3. `Download ZIP` 클릭
4. 압축 해제
5. 압축 해제한 폴더 사용


## 2. 폴더 구조

프로젝트 폴더는 아래 구조로 사용합니다.

```text
ConfirmationGenerator/
├─ data/
│  ├─ account/
│  │  └─ account.png
│  ├─ logo.png
│  ├─ stamp.png
│  └─ 일임계약 리스트_통합.xlsx
│
├─ pdf_customer/
│  └─ 고객용 PDF 저장 위치
│
├─ pdf_pb/
│  └─ PB용 PDF 저장 위치
│
├─ src/
│  ├─ config.py
│  ├─ increase_confirmation.py
│  ├─ decrease_confirmation.py
│  ├─ extension_confirmation.py
│  └─ termination_confirmation.py
│
├─ requirements.txt
└─ README.md
```



## 3. 필요한 파일 넣기

아래 파일들은 반드시 정해진 위치에 넣어야 합니다.

### 3-1. 통합 엑셀 파일

아래 위치에 엑셀 파일을 넣습니다.

```text
data/일임계약 리스트_통합.xlsx
```

파일명은 `src/config.py`에 적힌 이름과 같아야 합니다.

### 3-2. 회사 로고 이미지

아래 위치에 로고 파일을 넣습니다.

```text
data/logo.png
```

### 3-3. 회사 도장 이미지

아래 위치에 도장 파일을 넣습니다.

```text
data/stamp.png
```

### 3-4. 계좌 이미지

아래 위치에 계좌 이미지를 넣습니다.

```text
data/account/account.png
```

이 이미지는 확인서 PDF의 마지막 페이지에 첨부됩니다.



## 4. Python 패키지 설치

프로젝트 폴더에서 아래 명령어를 실행합니다.

```bash
pip install -r requirements.txt
```

설치가 끝나면 실행 준비가 완료됩니다.



## 5. 실행 방법

각 확인서별로 실행 파일이 다릅니다.

명령어는 프로젝트 폴더에서 실행합니다.

### 5-1. 증액확인서 생성

```bash
python src/increase_confirmation.py
```

입력값은 아래와 같습니다.

```text
계좌번호
증액금액
```

예시:

```text
계좌번호를 입력하세요: 1234567890-01
증액금액을 입력하세요: 10000000
```

### 5-2. 감액확인서 생성

```bash
python src/decrease_confirmation.py
```

입력값은 아래와 같습니다.

```text
계좌번호
평가금액
인출금액합계
자동이체여부(Y/N)
```

예시:

```text
계좌번호를 입력하세요: 1234567890-01
평가금액을 입력하세요: 150000000
인출금액합계를 입력하세요: 30000000
자동이체여부를 입력하세요(Y/N): Y
```

### 5-3. 계약 연장확인서 생성

```bash
python src/extension_confirmation.py
```

입력값은 아래와 같습니다.

```text
계좌번호
평가금액
자동이체여부(Y/N)
```

예시:

```text
계좌번호를 입력하세요: 1234567890-01
평가금액을 입력하세요: 150000000
자동이체여부를 입력하세요(Y/N): Y
```

### 5-4. 계약 해지확인서 생성

```bash
python src/termination_confirmation.py
```

입력값은 아래와 같습니다.

```text
계좌번호
평가금액
인출금액합계
자동이체여부(Y/N)
```

예시:

```text
계좌번호를 입력하세요: 1234567890-01
평가금액을 입력하세요: 150000000
인출금액합계를 입력하세요: 150000000
자동이체여부를 입력하세요(Y/N): N
```



## 6. PDF 저장 위치

생성된 PDF는 아래 두 폴더에 저장됩니다.

### 고객용 PDF

```text
pdf_customer/
```

고객에게 전달할 PDF입니다.

고객용 PDF 비밀번호는 고객 생년월일 6자리입니다.

예시:

```text
860104
```

### PB용 PDF

```text
pdf_pb/
```

내부 보관용 PDF입니다.

PB용 PDF 비밀번호는 생성일 기준 `yymmdd` 형식입니다.

예시:

```text
2026년 04월 29일 생성 → 260429
```



## 7. 생성 완료 화면

정상적으로 생성되면 터미널에 아래와 같이 출력됩니다.

```text
고객용 PDF 생성 완료: C:\ConfirmationGenerator\pdf_customer\파일명.pdf
고객용 PDF 비밀번호: 860104
고객 이메일: example@email.com

PB용 PDF 생성 완료: C:\ConfirmationGenerator\pdf_pb\파일명.pdf
PB용 PDF 비밀번호: 260429
```

출력된 경로에서 PDF 파일을 확인하면 됩니다.



## 8. 주요 기능

### 8-1. 계좌번호 기준 고객 정보 조회

계좌번호를 입력하면 엑셀에서 해당 고객 정보를 조회합니다.

상품별로 정보가 표시되는 행이 다를 수 있어, 프로그램이 실제 값이 있는 행을 찾아 사용합니다.

### 8-2. 로고 및 도장 자동 삽입

Word 문서 생성 후 회사 로고와 도장을 자동으로 넣습니다.

사용 파일은 아래와 같습니다.

```text
data/logo.png
data/stamp.png
```

도장은 회사명 문구 근처에 삽입됩니다.

### 8-3. 계좌 이미지 첨부

PDF 마지막 페이지에 계좌 이미지를 첨부합니다.

사용 파일은 아래와 같습니다.

```text
data/account/account.png
```

PDF 구성 예시는 아래와 같습니다.

```text
1페이지: 확인서
2페이지: 계좌 이미지
```

### 8-4. PDF 비밀번호 자동 설정

PDF는 고객용과 PB용으로 나뉘어 각각 비밀번호가 설정됩니다.

```text
고객용 PDF: 고객 생년월일 6자리
PB용 PDF: 생성일 yymmdd
```

예시:

```text
고객 생년월일: 1986-01-04 → 860104
생성일: 2026-04-29 → 260429
```

### 8-5. 고객 이메일 출력

PDF 생성 후 고객 이메일을 터미널에 함께 출력합니다.

고객 이메일은 엑셀에서 조회된 고객 정보 기준으로 가져옵니다.

---

## 9. 위치 조정

로고나 도장 위치를 조정해야 할 경우 각 Python 파일 상단의 값을 수정합니다.

예시:

```python
LOGO_LEFT_CM = 16.3
LOGO_TOP_CM = 0.5
LOGO_WIDTH_CM = 3.8
LOGO_HEIGHT_CM = 1.2

STAMP_OFFSET_X_CM = 13.8
STAMP_OFFSET_Y_CM = -1.9
STAMP_SIZE_CM = 2
```

### 로고 위치

```text
LOGO_LEFT_CM: 숫자가 커지면 오른쪽으로 이동
LOGO_TOP_CM: 숫자가 커지면 아래로 이동
LOGO_WIDTH_CM: 로고 가로 크기
LOGO_HEIGHT_CM: 로고 세로 크기
```

### 도장 위치

```text
STAMP_OFFSET_X_CM: 숫자가 커지면 오른쪽으로 이동
STAMP_OFFSET_Y_CM: 숫자가 작아지면 위로 이동
STAMP_SIZE_CM: 도장 크기
```

---

## 10. 자주 발생하는 문제

### 10-1. 패키지 오류가 나는 경우

아래와 같은 오류가 나올 수 있습니다.

```text
ModuleNotFoundError: No module named 'pikepdf'
```

필요한 패키지가 설치되지 않은 상태입니다.

아래 명령어를 실행합니다.

```bash
pip install -r requirements.txt
```

### 10-2. 엑셀 시트를 찾을 수 없는 경우

엑셀 시트명과 코드에 적힌 시트명이 다르면 오류가 발생합니다.

확인할 위치는 아래와 같습니다.

```text
src/config.py
각 확인서 Python 파일의 SHEET_NAME
```

엑셀 시트명과 코드의 시트명이 정확히 같아야 합니다.

띄어쓰기 하나만 달라도 오류가 날 수 있습니다.

### 10-3. PDF가 생성되지 않는 경우

아래 내용을 확인합니다.

```text
1. Excel 파일이 data 폴더에 있는지
2. Word와 Excel이 설치되어 있는지
3. logo.png, stamp.png, account.png가 정해진 위치에 있는지
4. pdf_customer, pdf_pb 폴더가 있는지
5. 입력한 계좌번호가 엑셀에 있는지
```

---

## 11. 사용 순서 요약

처음 사용하는 경우 아래 순서대로 진행하면 됩니다.

```text
1. 프로젝트 다운로드
2. 필요한 파일을 data 폴더에 넣기
3. pip install -r requirements.txt 실행
4. 필요한 확인서 Python 파일 실행
5. 입력값 입력
6. pdf_customer, pdf_pb 폴더에서 PDF 확인
```

---

## 12. 참고사항

- Windows 환경 기준으로 사용합니다.
- Microsoft Excel과 Word가 설치되어 있어야 합니다.
- 엑셀 파일명이나 이미지 파일명이 바뀌면 `src/config.py`도 같이 수정해야 합니다.
- PDF가 생성되지 않으면 먼저 `data` 폴더의 파일 위치를 확인합니다.

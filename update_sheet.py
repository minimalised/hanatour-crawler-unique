import os
import re
import json
import gspread
from google.oauth2.service_account import Credentials

# 1. Secrets에서 구글 계정 키 및 시트 ID 가져오기
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
google_json_str = os.environ["GOOGLE_JSON_RAW"]
service_account_info = json.loads(google_json_str)

credentials = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
gc = gspread.authorize(credentials)

# 스프레드시트 열기
spreadsheet_id = os.environ["SPREADSHEET_ID"]
doc = gc.open_by_key(spreadsheet_id)

# 2. 원본 시트('호텔상품리스트') 데이터 가져오기
source_sheet = doc.worksheet('호텔상품리스트')
all_rows = source_sheet.get_all_values()

if not all_rows:
    print("데이터가 없습니다.")
    exit(0)

# 3. 지정된 6개 열 가져오기 및 B열 괄호 제거
# 추출 대상 열 인덱스 (0부터 시작): A(0), B(1), C(2), D(3), F(5), H(7)
extracted_data = []

# 대괄호 및 소괄호 [, ], (, ) 제거용 정규표현식
bracket_pattern = re.compile(r'[\[\]\(\)]')

for row in all_rows:
    def get_val(idx):
        return row[idx] if idx < len(row) else ''

    col_a = get_val(0)  # id
    col_b = get_val(1)  # 상품명
    col_c = get_val(2)  # 가격
    col_d = get_val(3)  # 혜택가격
    col_f = get_val(5)  # 링크
    col_h = get_val(7)  # 이미지링크

    # B열 상품명에서 [, ], (, ) 제거
    cleaned_b = bracket_pattern.sub('', col_b)

    extracted_data.append([col_a, cleaned_b, col_c, col_d, col_f, col_h])

# 4. Target 시트('호텔github')에 데이터 반영
try:
    target_sheet = doc.worksheet('호텔github')
except gspread.exceptions.WorksheetNotFound:
    target_sheet = doc.add_worksheet(title='호텔github', rows="1000", cols="10")

# 기존 내용 덮어쓰기
target_sheet.clear()
target_sheet.update('A1', extracted_data)

print(f"총 {len(extracted_data)}행 데이터 업데이트 완료!")

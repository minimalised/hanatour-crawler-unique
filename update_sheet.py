import os
import re
import json
import gspread
from urllib.parse import urlparse, urlunparse
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

# 3. 지정된 6개 열 가져오기, B열 괄호 제거, F열 도메인 변경
extracted_data = []

# 대괄호 및 소괄호 [, ], (, ) 제거용 정규표현식
bracket_pattern = re.compile(r'[\[\]\(\)]')

# 도메인을 변경하는 함수
def replace_domain(url, new_domain="https://unique.hanatour.com"):
    if not url or not url.startswith("http"):
        return url
    
    parsed_url = urlparse(url)
    target_parsed = urlparse(new_domain)
    
    # scheme(https) 및 netloc(unique.hanatour.com) 변경 후 재조합
    updated_url = urlunparse((
        target_parsed.scheme,
        target_parsed.netloc,
        parsed_url.path,
        parsed_url.params,
        parsed_url.query,
        parsed_url.fragment
    ))
    return updated_url

for idx, row in enumerate(all_rows):
    def get_val(i):
        return row[i] if i < len(row) else ''

    col_a = get_val(0)  # id
    col_b = get_val(1)  # 상품명
    col_c = get_val(2)  # 가격
    col_d = get_val(3)  # 혜택가격
    col_f = get_val(5)  # 링크
    col_h = get_val(7)  # 이미지링크

    # 1행(헤더)은 변경 없이 그대로 유지
    if idx == 0:
        extracted_data.append([col_a, col_b, col_c, col_d, col_f, col_h])
        continue

    # B열 상품명에서 [, ], (, ) 제거
    cleaned_b = bracket_pattern.sub('', col_b)

    # F열 도메인을 https://unique.hanatour.com 으로 변경
    updated_f = replace_domain(col_f)

    extracted_data.append([col_a, cleaned_b, col_c, col_d, updated_f, col_h])

# 4. Target 시트('호텔github')에 데이터 반영
try:
    target_sheet = doc.worksheet('호텔github')
except gspread.exceptions.WorksheetNotFound:
    target_sheet = doc.add_worksheet(title='호텔github', rows="1000", cols="10")

# 기존 내용 덮어쓰기
target_sheet.clear()
target_sheet.update('A1', extracted_data)

print(f"총 {len(extracted_data)}행 데이터 업데이트 완료!")

#!/usr/bin/env python3
import concurrent.futures
import sys
import requests
import xlsxwriter
import openpyxl
import os
from io import BytesIO
from dataclasses import dataclass, field
from PIL import Image, ImageFont
import json
from typing import Optional
from enum import Enum, auto
from urllib.parse import urlparse, parse_qs
import yaml

VERSION = "1.0.2"

if not os.path.exists("config.yml"):
    with open("config.yml", "w", encoding="utf-8") as f:
        yaml.dump({
            "aladinKey": "write here",
            "libraryLink": "write here",
            "outputFileName": "output.xlsx"
        }, f, allow_unicode=True)

    print(f"config.yml 파일이 존재하지 않아 기본 템플릿을 생성했어요. 값을 채워넣고 다시 실행해주세요.")
    sys.exit(1)

with open("config.yml", encoding="utf-8") as f:
    config = yaml.safe_load(f)
    required_keys = ["aladinKey", "libraryLink", "outputFileName"]
    for k in required_keys:
        if k not in config or not config[k]:
            print(f"콘피그 파일에 '{k}' 값이 비어 있어요.")
            sys.exit(1)


query_params = parse_qs(urlparse(config["libraryLink"]).query)

SCHOOL_NAME = query_params.get('schoolName', [None])[0]
PROV_CODE = query_params.get('provCode', [None])[0]
NEIS_CODE = query_params.get('neisCode', [None])[0]

if SCHOOL_NAME is None or PROV_CODE is None or NEIS_CODE is None:
    print("올바르지 않은 도서관 링크가 제공되었어요. 프로그램의 설명을 참고하여 올바른 링크를 입력해주세요.")
    sys.exit(1)

print(f"@@@@@ 학교 정보를 로딩했어요: {SCHOOL_NAME} (교육청 코드: {PROV_CODE}, 나이스 코드: {NEIS_CODE})")

ALADIN_API_KEY = config["aladinKey"]
OUTPUT_XLSX_FILE = config["outputFileName"]

DEFAULT_FONT_SIZE_PT = 11
DESCRIPTION_WRAP_WIDTH = 25

INPUT_XLSX_FILE = "list.xlsx"

ALADIN_URL_TEMPLATE = (
    "http://www.aladin.co.kr/ttb/api/ItemLookUp.aspx"
    "?ttbkey={api_key}&itemIdType=ISBN&ItemId={isbn}"
    "&output=js&Version=20131101&OptResult=Story,categoryIdList,"
    "bestSellerRank,ratingInfo,reviewList"
)


class LibraryBookStatus(Enum):
    EXISTS = auto()
    NOT_EXISTS = auto()
    UNKNOWN = auto()

@dataclass
class Book:
    title: str
    link: str
    author: str
    publisher: str
    isbn13: str
    standard_price: int
    publish_date: str
    description: str
    rating_score: float
    rating_count: int
    category: str
    sheet_name: str
    cover: BytesIO = None
    library_status: LibraryBookStatus = LibraryBookStatus.UNKNOWN
    memo: str = ''
    book_key: int = None
    species_key: int = None

class Column:
    def __init__(self, header, getter, fmt_key):
        self.header = header
        self.getter = getter
        self.fmt_key = fmt_key

class FormatManager:
    def __init__(self, workbook, font_size_pt):
        base_args = {'font_size': font_size_pt, 'text_wrap': True}
        self.fmts = {
            'header': workbook.add_format({**base_args, 'bold': True, 'bg_color': '#D3D3D3', 'align': 'center', 'valign': 'vcenter'}),
            'center': workbook.add_format({**base_args, 'align': 'center', 'valign': 'vcenter'}),
            'left': workbook.add_format({**base_args, 'align': 'left', 'valign': 'vcenter'}),
            'price': workbook.add_format({**base_args, 'num_format': '#,##0"원"', 'align': 'center', 'valign': 'vcenter'}),
        }

    def get(self, key):
        return self.fmts.get(key)
    
COLUMNS = [
    Column('', lambda b: b.cover, None),
    Column('도서', lambda b: b.title, 'center'),
    Column('저자', lambda b: b.author, 'center'),
    Column('출판사', lambda b: b.publisher, 'center'),
    Column('ISBN13', lambda b: b.isbn13, 'center'),
    Column('정가', lambda b: b.standard_price, 'price'),
    Column('출판일', lambda b: b.publish_date, 'center'),
    Column('설명', lambda b: b.description, 'left'),
    Column('평점', lambda b: f'{b.rating_score:.1f} {"★"*round(b.rating_score/2)+"☆"*(5-round(b.rating_score/2))} ({b.rating_count})', 'center'),
    Column('카테고리', lambda b: b.category, 'left'),
    Column('교내 도서관 소장', lambda b: 'O' if b.library_status == LibraryBookStatus.EXISTS else ('X' if b.library_status == LibraryBookStatus.NOT_EXISTS else '?'), 'center'),
    Column('메모', lambda b: b.memo, 'left'),
]

def check_library(isbn: str, session: requests.Session, timeout: int = 10) -> LibraryBookStatus:
    url = "https://read365.edunet.net/alpasq/api/search"

    payload = {
        "searchKeyword": isbn,
        "neisCode": [NEIS_CODE],
        "provCode": PROV_CODE,
        "coverYn": "N"
    }

    headers = {"Content-Type": "application/json"}

    try:
        response = session.post(url, json=payload, headers=headers, timeout=timeout)
        response.raise_for_status()

        data = response.json()

        results = data.get("data", {}).get("bookList", [])

        for book in results:
            if book.get("isbn") == isbn:
                print(f"> [조회 성공][도서관] ISBN {isbn}: 소장하고 있는 도서에요 ㅡ bookKey: {book.get('bookKey')}, speciesKey: {book.get('speciesKey')}")
                return LibraryBookStatus.EXISTS, book.get('bookKey'), book.get('speciesKey')
            
        print(f"> [조회 성공][도서관] ISBN {isbn}: 소장하고 있지 않은 도서에요.")
        return LibraryBookStatus.NOT_EXISTS, None, None
    
    except requests.exceptions.HTTPError as http_err:
        print(f"> [조회 실패][도서관] ISBN {isbn}: HTTP 오류가 발생했어요 - {http_err.response.status_code} {http_err.response.reason}")
        return LibraryBookStatus.UNKNOWN, None, None
    
    except requests.exceptions.RequestException as e:
        print(f"> [조회 실패][도서관] ISBN {isbn}: 네트워크 오류가 발생했어요 - {e}")
        return LibraryBookStatus.UNKNOWN, None, None
    
    except Exception as e:
        print(f"> [조회 실패][도서관] ISBN {isbn}: 예상치 못한 오류가 발생했어요 - {e}")
        return LibraryBookStatus.UNKNOWN, None, None

def fetch(isbn: str, aladin_api_key: str, session: requests.Session, timeout: int = 5, memo: str = '', sheet_name: str = '') -> Optional[Book]:
    url = ALADIN_URL_TEMPLATE.format(api_key=ALADIN_API_KEY, isbn=isbn)
    book = None

    try:
        resp = session.get(url, timeout=timeout)
        resp.raise_for_status()

        items = resp.json().get("item", [])

        if not items:
            print(f"> [조회 실패][알라딘] ISBN {isbn}: 데이터를 찾을 수 없어요.")
            return None
        
        item = items[0]

        desc = item.get("description", "").strip()
        rating_info = item.get("subInfo", {}).get("ratingInfo", {})
        score = float(rating_info.get("ratingScore", 0))
        count = int(rating_info.get("ratingCount", 0))

        book = Book(
            title=item.get("title", ""),
            link=item.get("link", ""),
            author=item.get("author", ""),
            publisher=item.get("publisher", ""),
            isbn13=item.get("isbn13", ""),
            standard_price=int(item.get("priceStandard", 0)),
            publish_date=item.get("pubDate", ""),
            description=desc,
            rating_score=score,
            rating_count=count,
            category=item.get("categoryName", ""),
            sheet_name=sheet_name,
            memo=memo,
        )

        print(f"> [조회 성공][알라딘] ISBN {isbn}: '{book.title}' 정보를 가져왔어요.")

        if cover_url := item.get("cover"):
            try:
                cover_resp = session.get(cover_url, timeout=timeout)
                cover_resp.raise_for_status()

                book.cover = BytesIO(cover_resp.content)

            except requests.exceptions.RequestException as e:
                print(f"> [오류][알라딘] ISBN {isbn} 커버 이미지를 가져오는 중 오류가 발생했어요: {e}")
                book.cover_data = None

    except requests.exceptions.RequestException as e:
        print(f"> [조회 실패][알라딘] ISBN {isbn}: 정보 가져오는 중 오류가 발생했어요 - {e}")
        return None
    
    except json.JSONDecodeError:
        print(f"> [조회 실패][알라딘] ISBN {isbn}: 응답이 유효한 JSON 형식이 아니에요.")
        return None
    
    except Exception as e:
         print(f"> [조회 실패][알라딘] ISBN {isbn}: 정보 처리 중 예상치 못한 오류가 발생했어요 - {e}")
         return None
    
    if book:
        book.library_status, book.book_key, book.species_key = check_library(isbn, session)
        
    return book

def get_text_px(text: str, font: ImageFont.FreeTypeFont) -> tuple[int, int]:
    char_width_avg = font.getlength('A')
    char_height = font.size

    max_width = 0

    lines = text.split('\n')

    for line in lines:
        line_width = 0

        for char in line:
            line_width += char_width_avg * 2 if ord(char) > 127 else char_width_avg

        max_width = max(max_width, line_width)

    return int(max_width), int(len(lines) * char_height * 1.2)

def col_to_px(width):
    return int(width * 7 + 5)

def row_to_px(height):
    return int(height * 96 / 72)

def create(books: list[Book], output: str, font_size_pt: int):
    font = ImageFont.load_default()
    workbook = xlsxwriter.Workbook(output, {'default_date_format': 'yyyy-mm-dd'})
    fm = FormatManager(workbook, font_size_pt)

    # 입력 순서대로 시트명 목록 생성
    sheet_sequence = []
    for book in books:
        if book.sheet_name not in sheet_sequence:
            sheet_sequence.append(book.sheet_name)

    # 시트별로 작성
    for sheet_name in sheet_sequence:
        group_books = [book for book in books if book.sheet_name == sheet_name]
        worksheet = workbook.add_worksheet(sheet_name)
        img_idx = 0
        img_col_width_char = 16
        col_widths = [img_col_width_char] + [0] * (len(COLUMNS) - 1)
        avg_char_w = font.getlength('A')

        # 컬럼 너비 계산
        for idx, col in enumerate(COLUMNS):
            if idx == img_idx:
                continue
            header_w_px, _ = get_text_px(col.header, font)
            char_w = header_w_px / avg_char_w + 0.1
            col_widths[idx] = max(col_widths[idx], char_w)
            for book in group_books:
                val = col.getter(book)
                text = f'{int(val):,}원' if col.fmt_key == 'price' else str(val or '')
                w_px, _ = get_text_px(text, font)
                char_w = w_px / avg_char_w + 0.1
                col_widths[idx] = max(col_widths[idx], char_w)
            col_widths[idx] = min(col_widths[idx], 60)

        # 헤더 작성
        for idx, width in enumerate(col_widths):
            worksheet.set_column(idx, idx, width)
            worksheet.write(0, idx, COLUMNS[idx].header, fm.get('header'))

        # 행 높이 계산
        row_heights = [font_size_pt * 1.7]
        for book in group_books:
            max_h = font_size_pt * 1.7
            for idx, col in enumerate(COLUMNS):
                if idx == img_idx:
                    cell_h = 112
                else:
                    lines = (str(col.getter(book)) or '').count('\n') + 1
                    cell_h = lines * font_size_pt * 1.7
                max_h = max(max_h, cell_h)
            row_heights.append(max_h)

        for r, h in enumerate(row_heights):
            worksheet.set_row(r, h)

        # 데이터 작성
        for r, book in enumerate(group_books, start=1):
            for idx, col in enumerate(COLUMNS):
                if idx == img_idx:
                    continue
                val = col.getter(book)
                if col.fmt_key == 'price':
                    worksheet.write_number(r, idx, val, fm.get('price'))
                else:
                    worksheet.write(r, idx, val, fm.get(col.fmt_key or 'center'))

        # 이미지 삽입
        for r, book in enumerate(group_books, start=1):
            if book.cover:
                cell_w_px = col_to_px(col_widths[img_idx])
                cell_h_px = row_to_px(row_heights[r])
                im = Image.open(book.cover)
                buf = BytesIO()
                im_resized = im.resize((cell_w_px, cell_h_px), Image.Resampling.LANCZOS)
                im_resized.save(buf, format='PNG')
                buf.seek(0)
                worksheet.insert_image(r, img_idx, f"{book.isbn13}.png", {'image_data': buf, 'x_offset': 0, 'y_offset': 0, 'positioning': 1})

    workbook.close()
    print(f"@@@@@ 엑셀 파일({output})을 저장했어요.")

if __name__ == "__main__":
    books_to_check = []

    if not os.path.exists(INPUT_XLSX_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet['A1'] = 'ISBN13'
        sheet['B1'] = '시트'
        sheet['C1'] = '메모 (선택)'

        workbook.save(INPUT_XLSX_FILE)

        print(f"'{INPUT_XLSX_FILE}' 파일이 존재하지 않아 생성했어요. A열에 ISBN13, B열에 시트, C열에 메모를 입력해주세요.")
        sys.exit(1)

    try:
        workbook = openpyxl.load_workbook(INPUT_XLSX_FILE, data_only=True)
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, max_col=3):
            isbn_cell = row[0]
            sheet_cell = row[1] if len(row) > 1 else None
            memo_cell = row[2] if len(row) > 2 else None

            if isbn_cell.value:
                isbn = str(isbn_cell.value).strip()
                sheet_name = str(sheet_cell.value).strip() if sheet_cell and sheet_cell.value else ''
                memo = str(memo_cell.value).strip() if memo_cell and memo_cell.value else ''

                if not sheet_name:
                    print(f"'{isbn}' 책이 시트가 지정되지 않았어요.")
                    sys.exit(1)

                books_to_check.append((isbn, sheet_name, memo))
    except Exception as e:
        print(f"@@@@@ '{INPUT_XLSX_FILE}' 파일을 읽는 중 오류가 발생했어요: {e}")
        sys.exit(1)

    session = requests.Session()
    books = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(fetch, isbn, ALADIN_API_KEY, session, 5, memo, sheet_name)
                   for isbn, sheet_name, memo in books_to_check]
        for (isbn, sheet_name, memo), future in zip(books_to_check, futures):
            try:
                result = future.result()
                if result:
                    books.append(result)
            except Exception as exc:
                print(f"ISBN {isbn} 처리 중 예외가 발생했어요: {exc}")

    print(f"@@@@@ 총 {len(books)}권의 책 정보를 성공적으로 가져왔어요.")

    if not books:
        print("처리할 책 데이터가 없어요. 프로그램을 종료합니다.")
        sys.exit(1)

    create(books, OUTPUT_XLSX_FILE, DEFAULT_FONT_SIZE_PT)
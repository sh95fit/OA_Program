import os
import xlrd
from xlwt import Workbook
from xlutils.copy import copy
import re


def extract_year_month_from_filename(filename):
    pattern = r'(\d{2})년(\d{2})월'
    match = re.search(pattern, filename)

    if match:
        year = match.group(1)
        month = match.group(2)
        return year, month
    else:
        return None, None


def process_excel_file(file_path, new_workbook, year, month):
    try:
        # 엑셀 파일 열기
        excel_workbook = xlrd.open_workbook(file_path, formatting_info=True)
        sheet = excel_workbook.sheet_by_index(0)

        # 새로운 시트 생성
        new_sheet = new_workbook.add_sheet(sheet.name)

        # A열에서 값이 유효한 행 찾기
        for row_idx in range(0, sheet.nrows):
            for col_idx in range(sheet.ncols):
                new_sheet.write(row_idx, col_idx,
                                sheet.cell_value(row_idx, col_idx))

            if row_idx != 0:
                new_sheet.write(row_idx, 5, f'{year}')
                new_sheet.write(row_idx, 6, f'{month}')

        print(f"엑셀 파일 {file_path} 처리 완료")

    except Exception as e:
        print(f"엑셀 파일 {file_path} 처리 중 오류 발생: {e}")


def main():
    directory = r'C:\Users\user\Desktop\1차프로젝트\모니터링관련프로젝트\통합모니터링개발관련\기존KT클라우드서버\서버호스팅비용관리'

    new_workbook = Workbook()

    files = os.listdir(directory)

    xls_files = [f for f in files if f.endswith(
        '.xls') and 'servertotal' in f.lower()]

    for idx, xls_file in enumerate(xls_files):
        file_path = os.path.join(directory, xls_file)

        year, month = extract_year_month_from_filename(xls_file)

        if year and month:
            process_excel_file(file_path, new_workbook,
                               year, month)
        else:
            print(f"파일 이름에서 연도와 월을 추출할 수 없습니다: {xls_file}")

    new_workbook.save(os.path.join(directory, 'ServerHostingList.xls'))
    print("모든 파일 처리 완료 및 최종 파일 생성 완료")


if __name__ == "__main__":
    main()

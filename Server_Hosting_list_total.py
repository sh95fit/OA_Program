import os
import pandas as pd

# 디렉토리 경로
directory = r'C:\Users\user\Desktop\1차프로젝트\모니터링관련프로젝트\통합모니터링개발관련\기존KT클라우드서버\서버호스팅비용관리'

# 디렉토리 내 파일 목록 가져오기
file_list = [filename for filename in os.listdir(
    directory) if 'servertotal' in filename]

# 데이터를 저장할 DataFrame 생성
merged_df = pd.DataFrame()

# 파일들을 읽어서 DataFrame에 추가
for filename in file_list:
    # 파일 이름에서 YY와 MM 추출
    year = filename.split('_')[0][-2:]
    month = filename.split('_')[1][:2]

    # 엑셀 파일 읽기
    file_path = os.path.join(directory, filename)
    df = pd.read_excel(file_path)

    # 유효한 값이 있는 행 확인 후 YY와 MM 입력
    valid_rows = df.dropna(subset=[df.columns[0]])  # 첫 번째 열에 값이 있는 행 선택
    valid_rows['F'] = month  # 월 값은 F열에 입력
    valid_rows['G'] = year   # 연도 값은 G열에 입력

    # 데이터 프레임에 추가
    merged_df = pd.concat([merged_df, valid_rows])

# 새로운 엑셀 파일로 저장
output_file_path = os.path.join(directory, 'merged_servertotal.xlsx')

# 파일이 이미 존재하는 경우 덮어쓰기
if os.path.exists(output_file_path):
    os.remove(output_file_path)

merged_df.to_excel(output_file_path, index=False)

print("Merged file saved successfully at:", output_file_path)

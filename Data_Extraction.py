import os
import pandas as pd
import mysql.connector

from dotenv import load_dotenv

DB_HOST = os.getenv('db_host')
DB_PORT = os.getenv('db_port')
DB_USER = os.getenv('db_user')
DB_PASSWORD = os.getenv('db_password')
DB_DATABASE = os.getenv('db_database')

connection = mysql.connector.connect(
    host=DB_HOST,
    port=DB_PORT,
    user=DB_USER,
    password=DB_PASSWORD,
    database=DB_DATABASE,
)

rtu_id = 933
query = f"SELECT multi, cpg, save_time_id, save_time FROM gsmon_solar_data where rtu_id = {rtu_id}"

output_file = "C:/Users/user/Desktop/KIMSEHUN/develop/OA_Program/solar_data.csv"


def fetch_and_append_to_csv(connection, query, output_file, chunk_size=10000):
    offset = 0
    cnt = 0

    while True:
        df = pd.read_sql(
            query + f" LIMIT {chunk_size} OFFSET {offset}", connection)
        if df.empty:
            break

        offset += chunk_size

        df.to_csv(output_file, mode="a", index=True,
                  header=not os.path.exists(output_file))
        cnt += 1
        print(f"{cnt}회 batch 10000개 처리")


print("데이터 추출 진행중")
fetch_and_append_to_csv(connection, query, output_file)

connection.close()
print("데이터 추출 완료")

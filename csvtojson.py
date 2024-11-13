import csv
import json

def csv_to_json(csv_file_path, json_file_path):
    # CSV 파일 읽기
    with open(csv_file_path, 'r', encoding='utf-8') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        data = list(csv_reader)

    # JSON 파일로 쓰기
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

# 사용 예시
csv_file_path = '서희건설.csv'
json_file_path = '서희건설.json'
csv_to_json(csv_file_path, json_file_path)
print(f"CSV 파일 '{csv_file_path}'을(를) JSON 파일 '{json_file_path}'(으)로 변환했습니다.")
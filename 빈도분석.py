
import pandas as pd

# CSV 파일 불러오기
file_path = "빈도데이터.csv"  # 실제 CSV 파일 경로 입력
file_name = file_path.split(".")[0] # 파일명만 저장 

df = pd.read_csv(file_path)

# 컬럼명 공백 제거 (오류 방지)
df.columns = df.columns.str.strip()

# 빈도분석을 저장할 리스트 생성
freq_table = []

# 각 열별 빈도분석 수행
for col in df.columns:
    value_counts = df[col].value_counts()  # 빈도수
    total = len(df)  # 전체 개수
    first_row = True  # 첫 번째 행 여부 확인
    
    for category, count in value_counts.items():
        percent = round((count / total) * 100, 1)  # 퍼센트 계산
        
        if first_row:
            freq_table.append([col, category, count , f"{percent}%", f"{count} ({percent})%"])  # 첫 행에는 변수명 포함
            first_row = False  # 첫 번째 행 처리 완료
        else:
            freq_table.append(["", category, count, f"{percent}%", f"{count} ({percent})%"])  # 이후 행은 변수명 없이 처리

# DataFrame 변환
freq_df = pd.DataFrame(freq_table, columns=["Variables", "Categories", "n", "%", "n(%)"])

# 엑셀 파일로 저장
excel_file_path = f"excel_{file_name}.xlsx"
freq_df.to_excel(excel_file_path, index=False)






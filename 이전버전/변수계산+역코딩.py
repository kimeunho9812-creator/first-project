# main.py
import pandas as pd
from 역코딩 import reverse_coding
from 변수계산 import calculate_variables

# Excel 파일 경로 설정
file_path = r"C:\Users\82102\OneDrive\coing\매핑완료\매핑_데이터2.xlsx"
df = pd.read_excel(file_path, header=0)

# 1. 역코딩 수행
df, reverse_coded_variables = reverse_coding(df)

# 2. 변수 계산 수행 (여러 키워드를 처리)
target_keywords = input("변수 계산을 위한 키워드를 입력하세요 (여러 키워드는 콤마로 구분): ").strip()
target_keywords = [keyword.strip() for keyword in target_keywords.split(",")]

# 각 키워드에 대해 변수 계산
for keyword in target_keywords:
    # 계산에 포함할 변수 목록 생성
    calculation_vars = [
        var for var in df.columns 
        if keyword in var and (
            var in reverse_coded_variables or (keyword in var and var not in reverse_coded_variables)
        )
    ]
    
    # 원래 변수 제외, 역코딩된 변수 포함
    calculation_vars = [
        var for var in calculation_vars if not any(
            var.replace("_역코딩", "") == reverse_var.replace("_역코딩", "") for reverse_var in reverse_coded_variables
        ) or var in reverse_coded_variables
    ]
    
    # 선택된 변수 목록 출력
    print(f"'{keyword}' 키워드에 해당하는 변수들 (역코딩된 변수 포함, 원래 변수 제외): {calculation_vars}")

    # 각 행마다 선택된 변수의 합계와 평균을 계산하여 새로운 열에 추가
    if calculation_vars:
        df[f"{keyword}_합계"] = df[calculation_vars].sum(axis=1)
        df[f"{keyword}_평균"] = df[calculation_vars].mean(axis=1)
    else:
        print(f"'{keyword}' 키워드를 포함하는 유효한 변수가 없습니다.")

# 결과 저장
output_file_path = r"C:\Users\82102\OneDrive\coing\매핑완료\데이터파일_최종결과.xlsx"
df.to_excel(output_file_path, index=False)
print(f"모든 작업이 완료되었습니다. 결과가 {output_file_path}에 저장되었습니다.")

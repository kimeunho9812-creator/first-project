# 변수계산.py
import pandas as pd

def calculate_variables(df, reverse_coded_variables):
    # 키워드를 포함하면서 역코딩된 변수만 선택
    target_keyword = input("변수 계산을 위한 키워드를 입력하세요 (예: 조직웰빙): ").strip()
    calculation_vars = [var for var in reverse_coded_variables if target_keyword in var]

    # 선택된 변수 목록 출력
    print(f"'{target_keyword}' 키워드에 해당하는 역코딩된 변수들: {calculation_vars}")

    # 계산 수행
    if calculation_vars:
        df[f"{target_keyword}_합계"] = df[calculation_vars].sum(axis=1)
        df[f"{target_keyword}_평균"] = df[calculation_vars].mean(axis=1)
    else:
        print(f"'{target_keyword}' 키워드를 포함하는 역코딩된 변수가 없습니다.")
    
    return df

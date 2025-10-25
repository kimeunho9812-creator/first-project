# 역코딩.py
import pandas as pd

def reverse_coding(df):
    reverse_coded_variables = []
    
    # 변수 목록 출력
    print("Excel 파일의 변수 목록:")
    print(df.columns.tolist()) 

    # 역코딩 기능
    while True:
        # 여러 변수 이름을 콤마로 구분하여 입력
        target_variables = input("역코딩할 변수 이름을 입력하세요 (여러 변수는 콤마로 구분): ")
        target_variables = [var.strip() for var in target_variables.split(",")]
        
        # 선택된 변수의 최대값과 최소값 입력
        max_value = float(input("역코딩의 최대값을 입력하세요: "))
        min_value = float(input("역코딩의 최소값을 입력하세요: "))
        
        # 각 변수에 대해 역코딩 수행
        for var in target_variables:
            if var in df.columns:
                new_var = f"{var}_역코딩"
                df[new_var] = max_value + min_value - df[var]
                reverse_coded_variables.append(new_var)  
                print(f"'{var}' 변수에 대해 역코딩이 완료되었습니다.")
            else:
                print(f"변수 '{var}'은(는) Excel 파일에 존재하지 않습니다. 다시 확인해주세요.")
        
        # 추가 역코딩 여부 확인
        while True:
            more = input("최대값과 최소값이 다른 역코딩이 필요한 변수가 있습니까? (Y/N): ").strip().upper()
            if more in ["Y", "N"]:
                break
            print("잘못된 입력입니다. Y와 N 중 하나를 입력해주세요.")
        
        # 종료 조건
        if more == "N":
            break

         # 기존 변수 삭제
        original_variables = [var[0] for var in reverse_coded_variables]
        df.drop(columns=original_variables, inplace=True)
        print(f"기존 변수 {original_variables}가 삭제되었습니다.")
    
    return df, reverse_coded_variables

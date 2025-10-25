# !! 이 파일이 있는 폴더내에 작업을 수행할 파일을 추가하시고, file_path에서 그 파일명을 입력하시면 됩니다.


# 이 코드를 실행하면, 사용자가 input을 입력하여서 문항에 맞는 번호를 매핑하게 됩니다.
# 여러번 나오는 값은 한번만 입력하면 자동으로 매핑되며, 중복선택문항은 "|" 를 기준으로 값을 나누게 됩니다.
# 혹시 "|" 기호가 아니라 다른 값으로 중복선택문항을 만들었으면, ctrl + f 로 그 기호를 찾아서 수정하시면 됩니다.
# 입력이 전부 완료되고 나면, "매핑완료"라는 폴더에 두 개의 파일이 생깁니다.
# 하나는 문자열 에서 숫자로 모두 변환이 완료된 파일, 하나는 사용자가 입력한 매핑 정보를 담은 매핑정보 파일이 생깁니다.
# 사용자가 값을 오차없이 입력했는지 확인하기 위해, 매핑정보 파일을 확인해서 설문지의 번호와 맞게 매핑했는지 확인하시면 됩니다.

# p를 입력하면 그 값을 패스하게 됨. 
import pandas as pd
import os

# 파일 로드
file_path = "데이터2.xlsx"  # 파일명을 입력하세요
file_directory = os.path.dirname(os.path.abspath(file_path))  # 파일이 위치한 경로

folder_name = "매핑완료"  # 저장 폴더명
save_directory = os.path.join(file_directory, folder_name)  # 저장할 폴더 경로

# 폴더가 없으면 생성
if not os.path.exists(save_directory):
    os.makedirs(save_directory)
    print(f"✅ 폴더 생성 완료: {save_directory}")

# 시트 불러오기
xls = pd.ExcelFile(file_path)
df = xls.parse(xls.sheet_names[0])

# 데이터 정리 (공백 제거)
df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

# 변환된 데이터를 저장할 딕셔너리
user_defined_mappings = {}
global_mapping = {}  # 모든 컬럼에서 공통되는 매핑 저장
skip_values = set()  # 사용자가 'p'를 입력한 값 저장

# 사용자 입력을 통한 매핑 (같은 값은 한 번만 입력받음)
for col in df.columns:
    unique_values = set()
    
    for cell in df[col].dropna():
        if isinstance(cell, str):  # 문자열인 경우만 split 처리
            if cell.replace('.', '', 1).isdigit():
                continue
            unique_values.update(cell.split("|"))  
        elif isinstance(cell, (int, float)):  # 숫자는 변환하지 않음
            continue
    
    unique_values = sorted(unique_values)  # 정렬하여 일관성 유지

    # 숫자로 변환할 필요가 없는 경우 스킵
    if not unique_values:
        continue

    print(f"\n🔹 '{col}' 열의 고유 선택지: {unique_values}")
    mapping = {}

    # 사용자가 매핑할 숫자를 입력
    for value in unique_values:
        if value in global_mapping:
            mapping[value] = global_mapping[value]
        elif value in skip_values:  # 이전에 "p" 입력한 값이면 건너뜀
            mapping[value] = value  # 원본 값 유지
        else:
            while True:
                user_input = input(f"'{value}' → 숫자로 변환 (숫자 입력, 패스하려면 'p' 입력): ")
                if user_input.lower() == "p":
                    print(f"⚠ '{value}' 값은 패스됩니다. 원래 값 유지.")
                    mapping[value] = value  # 변환하지 않고 원래 값 유지
                    skip_values.add(value)  # 패스한 값 저장
                    break
                try:
                    mapping[value] = int(user_input)
                    global_mapping[value] = int(user_input)  # 전역 매핑 저장
                    break
                except ValueError:
                    print("⚠ 숫자로 입력하세요! (또는 'p' 입력)")

    # 변환 적용 (쉼표로 구분된 숫자로 변환, 패스한 값은 그대로 유지)
    df[col] = df[col].apply(lambda x: ",".join(map(str, [mapping[val] for val in x.split("|")])) 
                            if isinstance(x, str) and not x.replace('.', '', 1).isdigit() else x)
    
    user_defined_mappings[col] = mapping  # 컬럼별 매핑 저장

# 변환된 데이터 확인
print("\n✅ 변환 완료!")

# 매핑 정보를 데이터프레임으로 변환
mapping_df = pd.DataFrame([(col, key, value) for col, mapping in user_defined_mappings.items() for key, value in mapping.items()], 
                          columns=["컬럼명", "원본 값", "매핑된 값"])

mapping_df_unique = mapping_df.copy()
mapping_df_unique["컬럼명"] = mapping_df_unique["컬럼명"].mask(mapping_df_unique["컬럼명"].duplicated(), "")

# 변환된 파일 저장 경로
output_file_path = os.path.join(save_directory, f"매핑_{file_path}")
mapping_output_file_path = os.path.join(save_directory, "매핑정보.xlsx")

# 파일 저장
df.to_excel(output_file_path, index=False)
mapping_df_unique.to_excel(mapping_output_file_path, index=False)

print(f"✅ 변환된 파일 저장 완료: {output_file_path}")
print(f"✅ 매핑 정보 파일 저장 완료: {mapping_output_file_path}")

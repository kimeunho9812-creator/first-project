# 차이검정 자동화
# anova, 독립검정을 자동화하는 프로그램.

import numpy as np
import pandas as pd
import copy


# 같은 폴더 내에 파일이 있어야합니다. 파일을 바꾸고싶으시면 파일경로 변수에 파일을 변경하시면 됩니다
파일경로 = "3.차이검정.xlsx"
df = pd.read_excel(파일경로, header=None)

# 숫자형 데이터를 소수점 2자리까지 반올림, 2자리까지 표시
df = df.map(lambda x: f"{round(x,2):.2f}" if isinstance(x, float) else x)

# NaN값이 "nan" 문자열로 변한걸 다시 복구 

df = df.replace("nan", np.nan)




# 찾을 키워드 목록
keywords = ["집단통계량", "독립표본 검정", "기술통계", "ANOVA"]

# 결과 저장을 위한 딕셔너리
table_dict = {keyword: [] for keyword in keywords}

# 키워드별로 표를 자동으로 추출하여 저장
for keyword in keywords:
    # 키워드가 포함된 행의 인덱스 찾기
    start_indices = df[df.iloc[:, 0].astype(str).str.contains(keyword, na=False)].index.tolist()

    for start_idx in start_indices:
        # 표 끝을 찾기 위해 다음 행부터 검사
        for end_idx in range(start_idx + 1, len(df)):
            if df.iloc[end_idx].isnull().all():  # 빈 행이 나오면 종료
                break

        # 표 데이터 추출 및 저장
        table = df.iloc[start_idx:end_idx].dropna(how="all").reset_index(drop=True)
        table_dict[keyword].append(table)





# 개별 변수로 저장
집단통계량리스트 = table_dict["집단통계량"]
독립표본검정리스트 = table_dict["독립표본 검정"]
기술통계리스트 = table_dict["기술통계"]
ANOVA리스트 = table_dict["ANOVA"]



# 여기에는 없지만, 집단통계량이 2개면 집단통계량[1][0][1]
변수이름리스트 = []
for 집단통계량 in 집단통계량리스트:
    변수이름리스트.append(집단통계량[0][1])



# 여기서 실행한 코드로 종속변수들이 있는 열의 제목을 저장하게 된다.
unique_values = 집단통계량리스트[0][0].dropna().unique() # 집단통계량 0열의 유니크 값을 가져옴
# 여기서 "집단통계량", "성별"을 제외하면, 이건 종속변수들만 남는 리스트가 완성되는거다.
unique_values = unique_values[unique_values != "집단통계량"]
종속변수리스트 = unique_values[unique_values != 변수이름리스트[0]] # 집단통계량 0열 값중 "성별"을 없애기 위함.

# 종속변수리스트의 각 요소를 키로 하고, 빈 리스트를 값으로 가지는 딕셔너리 생성
종속변수딕셔너리 = {변수: [] for 변수 in 종속변수리스트}



# 남성, 여성 등 categories를 array형태로 저장할 리스트 생성
독립_카테고리_리스트 = []
기술_카테고리_리스트 = []

for 집단통계량 in 집단통계량리스트:
    독립_카테고리_리스트.append(집단통계량[1].dropna().unique())



# 마찬가지로 기술통계에서도 찾아서 append해준다.
for 기술통계 in 기술통계리스트:
    unique_values = 기술통계[1].dropna().unique()
    unique_values = unique_values[unique_values != "전체"]
    기술_카테고리_리스트.append(unique_values)

카테고리_리스트 = []
카테고리_리스트.append(독립_카테고리_리스트)
카테고리_리스트.append(기술_카테고리_리스트)
# np.array를 일반 리스트로 변환 (flatten)
flat_list = []
for sublist in 카테고리_리스트:
    for item in sublist:
        flat_list.extend(item.tolist())




# 딕셔너리에서 key를 찾은다음에, key에 해당하는 평균과 표준편차 값을 value에 추가

# 자녀가지각하는부모의학습관여를 집단통계량리스트[0]에서 찾는다.
# 행은같고, 3열에 평균, 4열에 표준편차가 있다.
# 이걸 남성, 여성이니까 2번 반복할건데, 이건 독립 카테고리 리스트의 길이를 기준으로 할거다.
# 만약 대도시,중소도시,읍면지역이면 3번을 돌려야한다.


for index,집단통계량 in enumerate(집단통계량리스트):
    for key, value in 종속변수딕셔너리.items():
        
        row, col = np.where(집단통계량.values == key)
        
        for i in range(len(독립_카테고리_리스트[index])):
            value.append(str(집단통계량.iloc[row[0] + i,3]) + "±" + str(집단통계량.iloc[row[0] + i,4]))




for index,기술통계 in enumerate(기술통계리스트):
    for key, value in 종속변수딕셔너리.items():
        print(key)
        row, col = np.where(기술통계.values == key)
        print(row,col)
        for i in range(len(기술_카테고리_리스트[index])):
            value.append(str(기술통계.iloc[row[0] + i,3]) + "±" + str(기술통계.iloc[row[0] + i,4]))





new_df_dict = {"Categories":flat_list}
for key, values in 종속변수딕셔너리.items():
    new_df_dict[key] = values  # key를 유지하고 값들만 추가

new_df = pd.DataFrame(new_df_dict)



# 1️⃣ 새로운 MultiIndex 열 이름 생성
new_columns = [("Categories", "")]  # 첫 번째 열 유지
for col in new_df.columns[1:]:  # 종속변수 열들
    new_columns.append((col, "M±SD"))
    new_columns.append((col, "t or F(p)"))  # 추가 열

# 2️⃣ 확장된 데이터 프레임 생성 (t or F(p) 값은 현재 없으므로 np.nan 추가)
expanded_data = []
for row in new_df.itertuples(index=False):
    new_row = [row[0]]  # Categories (예: 남학생, 여학생)
    for i in range(1, len(row)):  # 각 종속변수 컬럼 값
        new_row.append(row[i])  # M±SD 값
        new_row.append(np.nan)  # t or F(p) 값 (현재 없음)
    expanded_data.append(new_row)

# 3️⃣ MultiIndex 적용하여 DataFrame 생성
df_expanded = pd.DataFrame(expanded_data, columns=pd.MultiIndex.from_tuples(new_columns))



독립_t_결과리스트 = []
anova_결과리스트 = []



df_expanded.to_excel(f'F_차이_{파일경로}', index=True)
print("엑셀저장완료")



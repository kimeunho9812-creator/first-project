import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# GUI 설정
root = tk.Tk()
root.title("차이검정 자동화 프로그램")
root.geometry("400x250")

파일경로 = ""

# 파일 선택 함수
def 파일선택():
    global 파일경로
    파일경로 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if 파일경로:
        lbl_file.config(text=f"선택된 파일: {파일경로.split('/')[-1]}")

# 분석 실행 함수
def 분석실행():
    if not 파일경로:
        messagebox.showwarning("경고", "파일을 선택해주세요!")
        return

    try:
        df = pd.read_excel(파일경로, header=None)
        df = df.map(lambda x: f"{round(x,2):.2f}" if isinstance(x, float) else x)
        df = df.replace("nan", np.nan)

        keywords = ["집단통계량", "독립표본 검정", "기술통계", "ANOVA"]
        table_dict = {keyword: [] for keyword in keywords}

        for keyword in keywords:
            start_indices = df[df.iloc[:, 0].astype(str).str.contains(keyword, na=False)].index.tolist()
            for start_idx in start_indices:
                for end_idx in range(start_idx + 1, len(df)):
                    if df.iloc[end_idx].isnull().all():
                        break
                table = df.iloc[start_idx:end_idx].dropna(how="all").reset_index(drop=True)
                table_dict[keyword].append(table)

        집단통계량리스트 = table_dict["집단통계량"]
        기술통계리스트 = table_dict["기술통계"]

        unique_values = 집단통계량리스트[0][0].dropna().unique()
        unique_values = unique_values[unique_values != "집단통계량"]
        변수이름리스트 = [집단통계량[0][1] for 집단통계량 in 집단통계량리스트]
        종속변수리스트 = unique_values[unique_values != 변수이름리스트[0]]

        종속변수딕셔너리 = {변수: [] for 변수 in 종속변수리스트}
        독립_카테고리_리스트 = [list(집단통계량[1].dropna().unique()) for 집단통계량 in 집단통계량리스트]

        for index, 집단통계량 in enumerate(집단통계량리스트):
            for key, value in 종속변수딕셔너리.items():
                row, col = np.where(집단통계량.values == key)
                for i in range(len(독립_카테고리_리스트[index])):
                    value.append(str(집단통계량.iloc[row[0] + i, 3]) + "±" + str(집단통계량.iloc[row[0] + i, 4]))

        flat_list = []
        for sublist in 독립_카테고리_리스트:
            if isinstance(sublist, list):  # 리스트인지 확인 후 추가
                flat_list.extend(sublist)
            else:
                flat_list.append(sublist)  # 리스트가 아니라면 직접 추가

        new_df_dict = {"Categories": flat_list}
        for key, values in 종속변수딕셔너리.items():
            new_df_dict[key] = values

        new_df = pd.DataFrame(new_df_dict)

        new_columns = [("Categories", "")]
        for col in new_df.columns[1:]:
            new_columns.append((col, "M±SD"))
            new_columns.append((col, "t or F(p)"))

        expanded_data = []
        for row in new_df.itertuples(index=False):
            new_row = [row[0]]
            for i in range(1, len(row)):
                new_row.append(row[i])
                new_row.append(np.nan)
            expanded_data.append(new_row)

        df_expanded = pd.DataFrame(expanded_data, columns=pd.MultiIndex.from_tuples(new_columns))

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx;*.xls")],
                                                 title="분석 결과 저장 위치 선택",
                                                 initialfile="F_차이검정.xlsx")

        if save_path:
            df_expanded.to_excel(save_path, index=True)
            messagebox.showinfo("완료", f"분석이 완료되었습니다!\n파일 저장 위치: {save_path}")

    except Exception as e:
        messagebox.showerror("오류 발생", f"오류 내용: {str(e)}")


# GUI 구성
btn_select = tk.Button(root, text="파일 선택", command=파일선택)
btn_select.pack(pady=10)

lbl_file = tk.Label(root, text="선택된 파일: 없음", wraplength=350)
lbl_file.pack(pady=5)

btn_run = tk.Button(root, text="분석 실행", command=분석실행)
btn_run.pack(pady=10)

root.mainloop()

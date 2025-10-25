import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# 전역 변수
df = None
reverse_coded_variables = []


def select_file():
    """엑셀 파일 선택"""
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    
    if not file_path:
        return  # 사용자가 취소하면 그대로 종료
    
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, file_path)

    try:
        df = pd.read_excel(file_path)
        listbox_columns.delete(0, tk.END)  # 기존 목록 초기화
        for col in df.columns:
            listbox_columns.insert(tk.END, col)  # 컬럼명 리스트에 추가
    except Exception as e:
        messagebox.showerror("오류", f"파일을 열 수 없습니다: {e}")


def reverse_coding():
    """역코딩 수행"""
    global df, reverse_coded_variables

    if df is None:
        messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
        return
    
    selected_columns = [listbox_columns.get(idx) for idx in listbox_columns.curselection()]
    if not selected_columns:
        messagebox.showerror("오류", "역코딩할 변수를 선택하세요!")
        return

    try:
        max_value = float(entry_max_value.get())
        min_value = float(entry_min_value.get())
    except ValueError:
        messagebox.showerror("오류", "최대값과 최소값을 올바르게 입력하세요!")
        return

    reverse_coded_variables = []
    for var in selected_columns:
        new_var = f"{var}_역코딩"
        df[new_var] = max_value + min_value - df[var]
        reverse_coded_variables.append(new_var)

    # 기존 변수 삭제
    df.drop(columns=selected_columns, inplace=True)
    
    messagebox.showinfo("완료", f"역코딩 완료! 변수 {selected_columns}가 역코딩되었습니다.")
    update_result_text(f"역코딩 완료: {reverse_coded_variables}")


def calculate_variables():
    """변수 합계 및 평균 계산"""
    global df, reverse_coded_variables

    if df is None:
        messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
        return

    keyword = entry_keyword.get().strip()
    if not keyword:
        messagebox.showerror("오류", "변수 계산을 위한 키워드를 입력하세요!")
        return

    # 역코딩된 변수뿐만 아니라 원본 변수도 포함하도록 변경
    calculation_vars = [var for var in df.columns if keyword in var]

    if calculation_vars:
        df[f"{keyword}_합계"] = df[calculation_vars].sum(axis=1)
        df[f"{keyword}_평균"] = df[calculation_vars].mean(axis=1)
        messagebox.showinfo("완료", f"{keyword} 합계 및 평균 계산 완료!")
        update_result_text(f"{keyword}_합계 및 {keyword}_평균 계산 완료!")
    else:
        messagebox.showerror("오류", f"'{keyword}'를 포함하는 변수가 없습니다.")


def save_to_excel():
    """결과를 엑셀 파일로 저장"""
    if df is None:
        messagebox.showerror("오류", "먼저 분석을 실행하세요!")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return  # 사용자가 취소하면 그대로 종료

    try:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("저장 완료", f"결과가 {save_path}에 저장되었습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{e}")


def update_result_text(text):
    """결과 텍스트 업데이트"""
    text_result.delete(1.0, tk.END)
    text_result.insert(tk.END, text)


# Tkinter GUI 설정
root = tk.Tk()
root.title("역코딩 및 변수 계산 프로그램")

# 엑셀 파일 선택
frame_file = tk.Frame(root)
frame_file.pack(pady=5)
tk.Label(frame_file, text="엑셀 파일 경로:").pack(side=tk.LEFT)
entry_file_path = tk.Entry(frame_file, width=50)
entry_file_path.pack(side=tk.LEFT, padx=5)
btn_browse = tk.Button(frame_file, text="파일 선택", command=select_file)
btn_browse.pack(side=tk.LEFT)

# 변수 선택 (리스트 박스)
frame_columns = tk.Frame(root)
frame_columns.pack(pady=5)
tk.Label(frame_columns, text="역코딩할 변수 선택 (Ctrl+클릭으로 다중 선택 가능):").pack()
listbox_columns = tk.Listbox(frame_columns, width=50, height=10, selectmode=tk.MULTIPLE)
listbox_columns.pack()

# 역코딩 설정
frame_coding = tk.Frame(root)
frame_coding.pack(pady=5)
tk.Label(frame_coding, text="최대값:").pack(side=tk.LEFT)
entry_max_value = tk.Entry(frame_coding, width=10)
entry_max_value.pack(side=tk.LEFT, padx=5)
tk.Label(frame_coding, text="최소값:").pack(side=tk.LEFT)
entry_min_value = tk.Entry(frame_coding, width=10)
entry_min_value.pack(side=tk.LEFT, padx=5)

# 역코딩 실행 버튼
btn_reverse_coding = tk.Button(root, text="역코딩 실행", command=reverse_coding, bg="blue", fg="white")
btn_reverse_coding.pack(pady=5)

# 변수 계산 설정
frame_keyword = tk.Frame(root)
frame_keyword.pack(pady=5)
tk.Label(frame_keyword, text="변수 계산 키워드 입력 (예: 조직웰빙):").pack()
entry_keyword = tk.Entry(frame_keyword, width=30)
entry_keyword.pack()

# 변수 계산 실행 버튼
btn_calculate = tk.Button(root, text="변수 계산 (합계 & 평균)", command=calculate_variables, bg="green", fg="white")
btn_calculate.pack(pady=5)

# 결과 저장 버튼
btn_save = tk.Button(root, text="결과 저장", command=save_to_excel, bg="orange", fg="white")
btn_save.pack(pady=5)

# 결과 표시 창
frame_result = tk.Frame(root)
frame_result.pack(pady=5)
tk.Label(frame_result, text="결과:").pack(anchor="w")
text_result = tk.Text(frame_result, width=80, height=5)
text_result.pack()

# GUI 실행
root.mainloop()

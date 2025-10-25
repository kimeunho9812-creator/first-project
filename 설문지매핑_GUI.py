import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# 전역 변수
df = None
file_path = None
user_defined_mappings = {}
global_mapping = {}  # 전역 매핑 저장
skip_values = set()  # 'p'로 패스한 값 저장


def select_file():
    """엑셀 파일 선택"""
    global df, file_path
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
        messagebox.showinfo("파일 로드 완료", "엑셀 파일이 성공적으로 로드되었습니다!")
    except Exception as e:
        messagebox.showerror("오류", f"파일을 열 수 없습니다: {e}")


def perform_mapping():
    """사용자 입력을 받아 매핑 수행 (Shift 선택된 문항끼리는 같은 매핑 유지, 다른 문항과는 독립적으로 매핑)"""
    global df, user_defined_mappings, skip_values

    if df is None:
        messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
        return

    selected_columns = [listbox_columns.get(idx) for idx in listbox_columns.curselection()]
    if not selected_columns:
        messagebox.showerror("오류", "매핑할 변수를 선택하세요!")
        return

    shared_mapping = {}  # Shift 선택된 변수끼리 공유할 매핑 정보

    for col in selected_columns:
        unique_values = set()
        
        for cell in df[col].dropna():
            if isinstance(cell, str):  # 문자열인 경우만 split 처리
                if cell.replace('.', '', 1).isdigit():
                    continue
                unique_values.update(cell.split("|"))  
            elif isinstance(cell, (int, float)):  # 숫자는 변환하지 않음
                continue
        
        unique_values = sorted(unique_values)  # 정렬하여 일관성 유지

        if not unique_values:
            continue

        mapping = {}

        for value in unique_values:
            if col in user_defined_mappings and value in user_defined_mappings[col]:  
                # 이미 해당 문항(변수)에서 매핑한 값이 있으면 그대로 사용
                mapping[value] = user_defined_mappings[col][value]
            elif value in shared_mapping:  
                # Shift 선택된 다른 문항에서 매핑한 값이 있으면 동일한 값 사용
                mapping[value] = shared_mapping[value]
            elif value in skip_values:  
                # 'p' 입력한 값이면 건너뜀
                mapping[value] = value
            else:
                user_input = simpledialog.askstring(
                    "매핑 입력",
                    f"'{col}'에서 '{value}' → 숫자로 변환 (숫자 입력, 패스하려면 'p' 입력):"
                )
                if user_input is None:
                    return  # 사용자가 취소하면 중단

                if user_input.lower() == "p":
                    mapping[value] = value
                    skip_values.add(value)  # 패스한 값 저장
                else:
                    try:
                        mapping[value] = int(user_input)
                        shared_mapping[value] = int(user_input)  # 같은 그룹(Shift 선택된 문항)에서 공유
                    except ValueError:
                        messagebox.showerror("오류", "숫자로 입력하세요! (또는 'p' 입력)")
                        return

        # 선택한 변수만 매핑 적용
        df[col] = df[col].apply(
            lambda x: ",".join(map(str, [mapping[val] for val in x.split("|")])) 
            if isinstance(x, str) and not x.replace('.', '', 1).isdigit() else x
        )

        user_defined_mappings[col] = mapping  # 선택한 문항의 매핑 저장

    messagebox.showinfo("완료", "매핑이 완료되었습니다!")
    update_result_text("매핑 완료! 변환된 데이터를 저장하세요.")



def save_to_excel():
    """변환된 데이터를 엑셀 파일로 저장"""
    if df is None:
        messagebox.showerror("오류", "먼저 매핑을 실행하세요!")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return  # 사용자가 취소하면 그대로 종료

    try:
        # 매핑 정보를 데이터프레임으로 변환
        mapping_df = pd.DataFrame([(col, key, value) for col, mapping in user_defined_mappings.items() for key, value in mapping.items()], 
                                  columns=["컬럼명", "원본 값", "매핑된 값"])
        mapping_df_unique = mapping_df.copy()
        mapping_df_unique["컬럼명"] = mapping_df_unique["컬럼명"].mask(mapping_df_unique["컬럼명"].duplicated(), "")

        # 저장할 폴더 생성
        save_directory = os.path.dirname(save_path)
        mapping_output_file_path = os.path.join(save_directory, "매핑정보.xlsx")

        # 변환된 파일 저장
        df.to_excel(save_path, index=False)
        mapping_df_unique.to_excel(mapping_output_file_path, index=False)

        messagebox.showinfo("저장 완료", f"결과가 저장되었습니다:\n{save_path}\n매핑 정보: {mapping_output_file_path}")
    except Exception as e:
        messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{e}")


def update_result_text(text):
    """결과 텍스트 업데이트"""
    text_result.delete(1.0, tk.END)
    text_result.insert(tk.END, text)


# Tkinter GUI 설정
root = tk.Tk()
root.title("엑셀 데이터 매핑 프로그램")

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
tk.Label(frame_columns, text="매핑할 변수 선택 (Shift+클릭으로 연속 선택 가능, Ctrl+클릭으로 개별 선택 가능):").pack()
listbox_columns = tk.Listbox(frame_columns, width=50, height=10, selectmode=tk.EXTENDED)
listbox_columns.pack()

# 매핑 실행 버튼
btn_mapping = tk.Button(root, text="매핑 실행", command=perform_mapping, bg="blue", fg="white")
btn_mapping.pack(pady=5)

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

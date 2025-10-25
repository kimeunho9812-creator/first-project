import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import re

# 신뢰도 계산 결과 저장
results_log = []

def cronbach_alpha(data):
    """
    크론바흐 알파 계산 함수
    data: pandas DataFrame (선택된 문항들)
    """
    n_items = data.shape[1]
    item_variances = data.var(axis=0, ddof=1)
    total_variance = data.sum(axis=1).var(ddof=1)
    alpha = (n_items / (n_items - 1)) * (1 - (item_variances.sum() / total_variance))
    return alpha

def select_file():
    """엑셀 파일 선택"""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, file_path)

    # 파일 읽기 및 문항명 로드
    try:
        global df
        df = pd.read_excel(file_path)
        global column_names
        column_names = list(df.columns)
        update_recommendations()  # 전체 문항 표시
    except Exception as e:
        messagebox.showerror("오류", f"파일을 열 수 없습니다: {e}")

def expand_columns(selected_columns):
    """
    범위 입력 처리 (e.g., '희망1 to 희망6')
    """
    expanded_columns = []
    for col in selected_columns:
        match = re.match(r"(.+)(\d+)\s*to\s*(.+)(\d+)", col.strip(), re.IGNORECASE)
        if match:
            prefix_start, start, prefix_end, end = match.groups()
            if prefix_start == prefix_end:  # 같은 prefix일 경우
                expanded_columns.extend([f"{prefix_start}{i}" for i in range(int(start), int(end) + 1)])
            else:
                messagebox.showerror("오류", "범위 입력의 시작과 끝 문항명이 일치해야 합니다!")
        else:
            expanded_columns.append(col.strip())  # 범위가 아닌 경우 그대로 추가
    return expanded_columns

def update_recommendations():
    """전체 문항 표시"""
    listbox_recommendations.delete(0, tk.END)
    for col in column_names:
        listbox_recommendations.insert(tk.END, col)

def add_selected_recommendation(event):
    """추천 문항을 선택하여 추가"""
    selected = listbox_recommendations.get(listbox_recommendations.curselection())
    current_text = entry_columns.get()
    if current_text:
        entry_columns.insert(tk.END, f",{selected}")
    else:
        entry_columns.insert(tk.END, selected)

def calculate_alpha():
    """신뢰도 분석 및 결과 계산"""
    file_path = entry_file_path.get()
    selected_columns = entry_columns.get().strip()

    if not file_path:
        messagebox.showerror("오류", "엑셀 파일을 선택하세요!")
        return

    if not selected_columns:
        messagebox.showerror("오류", "문항명을 입력하세요!")
        return

    try:
        # 입력된 문항명 처리
        raw_columns = [col.strip() for col in selected_columns.split(",")]
        
        # 중복된 문항 체크
        if len(raw_columns) != len(set(raw_columns)):
            messagebox.showerror("오류", "동일한 문항이 두 번 이상 들어갔습니다.")
            return
        
        # 범위 확장 처리
        columns = expand_columns(raw_columns)
        data_for_alpha = df[columns]

        # 크론바흐 알파 계산
        alpha_value = cronbach_alpha(data_for_alpha)

        # "문항명"에서 숫자 제거 (e.g., '희망1 to 희망6' -> '희망')
        first_column = raw_columns[0].strip()
        base_name = ''.join(filter(str.isalpha, first_column.split()[0]))  # 숫자 및 범위 제거

        # 문항 삭제 시 크론바흐 알파 계산
        removed_alpha_values = {}
        for col in columns:
            remaining_data = data_for_alpha.drop(columns=[col])
            removed_alpha_values[col] = cronbach_alpha(remaining_data)

        # 결과 로그 저장
        results_log.append({
            "문항명": base_name,
            "문항 수": len(columns),
            "Cronbach’s α": round(alpha_value, 3),  # 소수점 세 자리로 반올림
            "문항 제거 시 알파 값": {k: round(v, 3) for k, v in removed_alpha_values.items()}  # 소수점 반올림
        })
        update_results_log()

        # 결과 표시
        result_text = f"Cronbach’s α: {round(alpha_value, 3)}\n\n"
        result_text += "각 문항 제거 시 Cronbach’s α:\n"
        for col, value in removed_alpha_values.items():
            result_text += f"{col} 제거 시 α: {round(value, 3)}\n"

        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)

    except Exception as e:
        messagebox.showerror("오류", f"분석 중 오류가 발생했습니다:\n{e}")

def update_results_log():
    """결과 로그 업데이트"""
    text_log.delete(1.0, tk.END)
    for i, result in enumerate(results_log, 1):
        text_log.insert(tk.END, f"[{i}] 변수: {result['문항명']}\n")
        text_log.insert(tk.END, f"    문항 수: {result['문항 수']}\n")
        text_log.insert(tk.END, f"    Cronbach’s α: {result['Cronbach’s α']:.3f}\n")
        text_log.insert(tk.END, f"    문항 제거 시 Cronbach’s α:\n")
        for col, value in result['문항 제거 시 알파 값'].items():
            text_log.insert(tk.END, f"        {col}: {value:.3f}\n")
        text_log.insert(tk.END, "\n")

def save_results_to_excel_custom():
    """결과를 사용자 정의 배치로 엑셀 파일에 저장"""
    if not results_log:
        messagebox.showinfo("정보", "저장할 결과가 없습니다.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return

    try:
        # 사용자 정의 배치 데이터 생성
        rows = []
        for result in results_log:
            # 전체 Cronbach’s α
            rows.append({
                "변수": result["문항명"],  # "문항명"을 "변수"로 변경
                "문항 수": result["문항 수"],
                "Cronbach’s α": result["Cronbach’s α"]
            })

        # DataFrame 생성
        df_results = pd.DataFrame(rows)

        # 엑셀로 저장
        df_results.to_excel(save_path, index=False)
        messagebox.showinfo("성공", f"결과가 {save_path}에 저장되었습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"결과 저장 중 오류가 발생했습니다:\n{e}")


# Tkinter GUI 설정
root = tk.Tk()
root.title("신뢰도 분석 (크론바흐 알파)")

# 엑셀 파일 선택
frame_file = tk.Frame(root)
frame_file.pack(pady=5)
tk.Label(frame_file, text="엑셀 파일 경로:").pack(side=tk.LEFT)
entry_file_path = tk.Entry(frame_file, width=50)
entry_file_path.pack(side=tk.LEFT, padx=5)
btn_browse = tk.Button(frame_file, text="찾아보기", command=select_file)
btn_browse.pack(side=tk.LEFT)

# 문항 입력
frame_columns = tk.Frame(root)
frame_columns.pack(pady=5)
tk.Label(frame_columns, text="문항명 (쉼표로 구분, 범위 입력 지원):").pack(side=tk.LEFT)
entry_columns = tk.Entry(frame_columns, width=50)
entry_columns.pack(side=tk.LEFT, padx=5)

# 추천 문항 표시
frame_recommendations = tk.Frame(root)
frame_recommendations.pack(pady=5)
tk.Label(frame_recommendations, text="문항 리스트:").pack(anchor="w")
listbox_recommendations = tk.Listbox(frame_recommendations, width=50, height=10)
listbox_recommendations.pack()
listbox_recommendations.bind("<Double-Button-1>", add_selected_recommendation)

# 분석 버튼
btn_analyze = tk.Button(root, text="분석 실행", command=calculate_alpha, bg="blue", fg="white")
btn_analyze.pack(pady=10)

# 결과 저장 버튼
btn_save = tk.Button(root, text="결과 저장", command=save_results_to_excel_custom, bg="green", fg="white")
btn_save.pack(pady=10)

# 결과 표시
frame_result = tk.Frame(root)
frame_result.pack(pady=5)
tk.Label(frame_result, text="결과:").pack(anchor="w")
text_result = tk.Text(frame_result, width=80, height=5)
text_result.pack()

# 결과 로그
frame_log = tk.Frame(root)
frame_log.pack(pady=5)
tk.Label(frame_log, text="결과 로그:").pack(anchor="w")
text_log = tk.Text(frame_log, width=80, height=10, bg="#f0f0f0")
text_log.pack()

# GUI 실행
root.mainloop()
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import re

# 신뢰도 계산 결과 저장
results_log = []

def cronbach_alpha(data):
    """크론바흐 알파 계산 함수"""
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

    try:
        global df
        df = pd.read_excel(file_path)
        global column_names
        column_names = list(df.columns)
        update_recommendations()
    except Exception as e:
        messagebox.showerror("오류", f"파일을 열 수 없습니다: {e}")

def expand_columns(selected_columns):
    """범위 입력 처리 (e.g., '희망1 to 희망6')"""
    expanded_columns = []
    for col in selected_columns:
        match = re.match(r"(.+)(\d+)\s*to\s*(.+)(\d+)", col.strip(), re.IGNORECASE)
        if match:
            prefix_start, start, prefix_end, end = match.groups()
            if prefix_start == prefix_end:
                expanded_columns.extend([f"{prefix_start}{i}" for i in range(int(start), int(end) + 1)])
            else:
                messagebox.showerror("오류", "범위 입력의 시작과 끝 문항명이 일치해야 합니다!")
        else:
            expanded_columns.append(col.strip())
    return expanded_columns

def update_recommendations():
    """전체 문항 표시"""
    listbox_recommendations.delete(0, tk.END)
    for col in column_names:
        listbox_recommendations.insert(tk.END, col)

def add_multiple_selected_recommendations(event):
    """Shift 키를 이용한 다중 선택 추가"""
    selected_items = listbox_recommendations.curselection()
    selected_columns = [listbox_recommendations.get(i) for i in selected_items]

    current_text = entry_columns.get()
    new_text = ", ".join(selected_columns)

    if current_text:
        entry_columns.insert(tk.END, f", {new_text}")
    else:
        entry_columns.insert(tk.END, new_text)

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
        raw_columns = [col.strip() for col in selected_columns.split(",")]
        if len(raw_columns) != len(set(raw_columns)):
            messagebox.showerror("오류", "동일한 문항이 두 번 이상 들어갔습니다.")
            return

        columns = expand_columns(raw_columns)
        data_for_alpha = df[columns]
        alpha_value = cronbach_alpha(data_for_alpha)

        first_column = raw_columns[0].strip()
        base_name = ''.join(filter(str.isalpha, first_column.split()[0]))

        removed_alpha_values = {}
        for col in columns:
            remaining_data = data_for_alpha.drop(columns=[col])
            removed_alpha_values[col] = cronbach_alpha(remaining_data)

        results_log.append({
            "문항명": base_name,
            "문항 수": len(columns),
            "Cronbach_alpha": round(alpha_value, 3),
            "문항 제거 시 알파 값": {k: round(v, 3) for k, v in removed_alpha_values.items()}
        })
        update_results_log()

        result_text = f"Cronbach’s α: {round(alpha_value, 3)}\n\n"
        result_text += "각 문항 제거 시 Cronbach’s α:\n"
        for col, value in removed_alpha_values.items():
            result_text += f"{col} 제거 시 α: {round(value, 3)}\n"

        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, result_text)

        entry_columns.delete(0, tk.END)

    except Exception as e:
        messagebox.showerror("오류", f"분석 중 오류가 발생했습니다:\n{e}")

def update_results_log():
    """결과 로그 업데이트"""
    text_log.delete(1.0, tk.END)
    for i, result in enumerate(results_log, 1):
        text_log.insert(tk.END, f"[{i}] 변수: {result['문항명']}\n")
        text_log.insert(tk.END, f"    문항 수: {result['문항 수']}\n")
        text_log.insert(tk.END, f"    Cronbach's α: {result['Cronbach_alpha']:.3f}\n")
        text_log.insert(tk.END, f"    문항 제거 시 Cronbach's α:\n")
        for col, value in result['문항 제거 시 알파 값'].items():
            text_log.insert(tk.END, f"        {col}: {value:.3f}\n")
        text_log.insert(tk.END, "\n")

def save_results_to_excel_custom():
    """결과를 엑셀 파일에 저장"""
    if not results_log:
        messagebox.showinfo("정보", "저장할 결과가 없습니다.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not save_path:
        return

    try:
        rows = []
        for result in results_log:
            # 기본 정보
            row_data = {
                "변수": result["문항명"],
                "문항 수": result["문항 수"],
                "Cronbach_alpha": result["Cronbach_alpha"]
            }

            # 각 문항 제거 시 알파 값을 추가
            for item_name, alpha_value in result["문항 제거 시 알파 값"].items():
                row_data[f"{item_name}_제거시"] = alpha_value

            rows.append(row_data)

        df_results = pd.DataFrame(rows)

        # 컬럼명을 보기 좋게 변경 (저장 직전)
        column_rename = {"Cronbach_alpha": "Cronbach's α"}
        for col in df_results.columns:
            if col.endswith("_제거시"):
                original_col = col.replace("_제거시", " 제거 시")
                column_rename[col] = original_col

        df_results.rename(columns=column_rename, inplace=True)
        df_results.to_excel(save_path, index=False, engine='openpyxl')
        messagebox.showinfo("성공", f"결과가 {save_path}에 저장되었습니다.")
    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        messagebox.showerror("오류", f"결과 저장 중 오류가 발생했습니다:\n{e}\n\n상세:\n{error_detail}")

# Tkinter GUI 설정
root = tk.Tk()
root.title("신뢰도 분석 (크론바흐 알파)")
root.geometry("900x850")
root.configure(bg="#f5f5f5")

# 색상 및 폰트 설정
COLOR_BG = "#f5f5f5"
COLOR_PRIMARY = "#2c3e50"
COLOR_SECONDARY = "#3498db"
COLOR_SUCCESS = "#27ae60"
COLOR_ACCENT = "#e74c3c"
COLOR_WHITE = "#ffffff"
COLOR_LIGHT_GRAY = "#ecf0f1"
COLOR_TEXT = "#2c3e50"

FONT_TITLE = ("맑은 고딕", 11, "bold")
FONT_NORMAL = ("맑은 고딕", 10)
FONT_SMALL = ("맑은 고딕", 9)

# 메인 컨테이너
main_container = tk.Frame(root, bg=COLOR_BG)
main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

# ==================== 파일 선택 섹션 ====================
file_frame = tk.LabelFrame(main_container, text=" 1. 데이터 파일 선택 ",
                           font=FONT_TITLE, bg=COLOR_WHITE, fg=COLOR_PRIMARY,
                           padx=15, pady=15, relief=tk.RIDGE, borderwidth=2)
file_frame.pack(fill=tk.X, pady=(0, 10))

tk.Label(file_frame, text="엑셀 파일:", font=FONT_NORMAL, bg=COLOR_WHITE, fg=COLOR_TEXT).grid(
    row=0, column=0, sticky="w", padx=(0, 10))
entry_file_path = tk.Entry(file_frame, width=60, font=FONT_NORMAL,
                           relief=tk.SOLID, borderwidth=1)
entry_file_path.grid(row=0, column=1, padx=(0, 10), ipady=5)
btn_browse = tk.Button(file_frame, text="📁 찾아보기", command=select_file,
                       font=FONT_NORMAL, bg=COLOR_SECONDARY, fg=COLOR_WHITE,
                       relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                       activebackground="#2980b9", activeforeground=COLOR_WHITE)
btn_browse.grid(row=0, column=2)

# ==================== 문항 입력 섹션 ====================
input_frame = tk.LabelFrame(main_container, text=" 2. 분석 문항 선택 ",
                            font=FONT_TITLE, bg=COLOR_WHITE, fg=COLOR_PRIMARY,
                            padx=15, pady=15, relief=tk.RIDGE, borderwidth=2)
input_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

# 문항명 입력
tk.Label(input_frame, text="선택된 문항:", font=FONT_NORMAL, bg=COLOR_WHITE, fg=COLOR_TEXT).pack(
    anchor="w", pady=(0, 5))
entry_columns = tk.Entry(input_frame, font=FONT_NORMAL, relief=tk.SOLID, borderwidth=1)
entry_columns.pack(fill=tk.X, ipady=5, pady=(0, 10))

tk.Label(input_frame, text="💡 Tip: 쉼표로 구분하거나, 범위 입력 지원 (예: 희망1 to 희망6)",
         font=FONT_SMALL, bg=COLOR_WHITE, fg="#7f8c8d").pack(anchor="w", pady=(0, 10))

# 문항 리스트 (스크롤바 포함)
tk.Label(input_frame, text="문항 리스트 (더블클릭으로 선택):", font=FONT_NORMAL,
         bg=COLOR_WHITE, fg=COLOR_TEXT).pack(anchor="w", pady=(0, 5))

listbox_frame = tk.Frame(input_frame, bg=COLOR_WHITE)
listbox_frame.pack(fill=tk.BOTH, expand=True)

scrollbar_list = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL)
scrollbar_list.pack(side=tk.RIGHT, fill=tk.Y)

listbox_recommendations = tk.Listbox(listbox_frame, font=FONT_NORMAL,
                                     selectmode=tk.EXTENDED, relief=tk.SOLID,
                                     borderwidth=1, yscrollcommand=scrollbar_list.set,
                                     selectbackground=COLOR_SECONDARY,
                                     selectforeground=COLOR_WHITE)
listbox_recommendations.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar_list.config(command=listbox_recommendations.yview)
listbox_recommendations.bind("<Double-Button-1>", add_multiple_selected_recommendations)

# ==================== 버튼 섹션 ====================
button_frame = tk.Frame(main_container, bg=COLOR_BG)
button_frame.pack(fill=tk.X, pady=(0, 10))

btn_analyze = tk.Button(button_frame, text="🔍 분석 실행", command=calculate_alpha,
                        font=("맑은 고딕", 11, "bold"), bg=COLOR_SECONDARY, fg=COLOR_WHITE,
                        relief=tk.FLAT, padx=30, pady=10, cursor="hand2",
                        activebackground="#2980b9", activeforeground=COLOR_WHITE)
btn_analyze.pack(side=tk.LEFT, padx=(0, 10))

btn_save = tk.Button(button_frame, text="💾 결과 저장", command=save_results_to_excel_custom,
                     font=("맑은 고딕", 11, "bold"), bg=COLOR_SUCCESS, fg=COLOR_WHITE,
                     relief=tk.FLAT, padx=30, pady=10, cursor="hand2",
                     activebackground="#229954", activeforeground=COLOR_WHITE)
btn_save.pack(side=tk.LEFT)

# ==================== 현재 분석 결과 섹션 ====================
result_frame = tk.LabelFrame(main_container, text=" 3. 현재 분석 결과 ",
                             font=FONT_TITLE, bg=COLOR_WHITE, fg=COLOR_PRIMARY,
                             padx=15, pady=15, relief=tk.RIDGE, borderwidth=2)
result_frame.pack(fill=tk.X, pady=(0, 10))

result_text_frame = tk.Frame(result_frame, bg=COLOR_WHITE)
result_text_frame.pack(fill=tk.BOTH, expand=True)

scrollbar_result = tk.Scrollbar(result_text_frame, orient=tk.VERTICAL)
scrollbar_result.pack(side=tk.RIGHT, fill=tk.Y)

text_result = tk.Text(result_text_frame, font=FONT_NORMAL, height=6,
                      relief=tk.SOLID, borderwidth=1, bg=COLOR_LIGHT_GRAY,
                      yscrollcommand=scrollbar_result.set, wrap=tk.WORD)
text_result.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar_result.config(command=text_result.yview)

# ==================== 전체 결과 로그 섹션 ====================
log_frame = tk.LabelFrame(main_container, text=" 4. 전체 결과 로그 ",
                          font=FONT_TITLE, bg=COLOR_WHITE, fg=COLOR_PRIMARY,
                          padx=15, pady=15, relief=tk.RIDGE, borderwidth=2)
log_frame.pack(fill=tk.BOTH, expand=True)

log_text_frame = tk.Frame(log_frame, bg=COLOR_WHITE)
log_text_frame.pack(fill=tk.BOTH, expand=True)

scrollbar_log = tk.Scrollbar(log_text_frame, orient=tk.VERTICAL)
scrollbar_log.pack(side=tk.RIGHT, fill=tk.Y)

text_log = tk.Text(log_text_frame, font=FONT_NORMAL,
                   relief=tk.SOLID, borderwidth=1, bg=COLOR_LIGHT_GRAY,
                   yscrollcommand=scrollbar_log.set, wrap=tk.WORD)
text_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar_log.config(command=text_log.yview)

root.mainloop()

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

# 전역 변수 안전 초기화
df = None
original_column_order = []
reverse_coded_columns = {}  # {원본컬럼: 역코딩컬럼} 매핑
selected_factor_vars = []
is_processing_click = False  # 더블클릭 처리 중인지 확인하는 플래그
original_file_path = None  # 원본 파일 경로 저장

# GUI 컴포넌트 전역 변수 (초기화 전 None으로 설정)
root = None
listbox_factor_vars = None
text_result = None
entry_file_path = None

# 버튼 전역 변수
btn_browse = None
btn_select_factor = None
btn_reverse = None
btn_calculate = None
btn_prepare_next = None
btn_save = None
btn_reset = None
btn_quick_calc = None  # 빠른 계산 버튼


def update_button_states():
    """워크플로우 상태에 따라 버튼 색상 및 활성화 상태 업데이트"""
    global workflow_state, btn_browse, btn_select_factor, btn_reverse, btn_calculate, btn_prepare_next

    try:
        # 버튼들이 존재하는지 확인
        if not all([btn_browse, btn_select_factor, btn_reverse, btn_calculate, btn_prepare_next]):
            print(f"버튼 상태: browse={btn_browse is not None}, select={btn_select_factor is not None}, reverse={btn_reverse is not None}, calc={btn_calculate is not None}, next={btn_prepare_next is not None}")
            return

        # 기본 색상으로 초기화
        btn_browse.config(bg=COLORS['primary'])
        btn_select_factor.config(bg=COLORS['primary'])
        btn_reverse.config(bg=COLORS['success'])
        btn_calculate.config(bg=COLORS['info'])
        btn_prepare_next.config(bg=COLORS['info'])

        # 단계별 상태 업데이트
        if workflow_state['step'] == 1:  # 파일 선택 단계
            btn_browse.config(bg=COLORS['active'], relief=tk.RAISED)
            btn_select_factor.config(state=tk.DISABLED, bg=COLORS['disabled'])
            btn_reverse.config(state=tk.DISABLED, bg=COLORS['disabled'])
            btn_calculate.config(state=tk.DISABLED, bg=COLORS['disabled'])
            btn_prepare_next.config(state=tk.DISABLED, bg=COLORS['disabled'])

        elif workflow_state['step'] == 2:  # 변수 선택 단계
            btn_browse.config(bg=COLORS['completed'])
            btn_select_factor.config(bg=COLORS['active'], relief=tk.RAISED, state=tk.NORMAL)
            btn_reverse.config(state=tk.DISABLED, bg=COLORS['disabled'])
            btn_calculate.config(state=tk.DISABLED, bg=COLORS['disabled'])
            btn_prepare_next.config(state=tk.DISABLED, bg=COLORS['disabled'])

        elif workflow_state['step'] == 3:  # 역코딩/계산 선택 단계
            btn_browse.config(bg=COLORS['completed'])
            btn_select_factor.config(bg=COLORS['completed'])
            btn_reverse.config(bg=COLORS['glow'], relief=tk.RAISED, state=tk.NORMAL)
            btn_calculate.config(bg=COLORS['glow'], relief=tk.RAISED, state=tk.NORMAL)
            btn_prepare_next.config(state=tk.DISABLED, bg=COLORS['disabled'])

        elif workflow_state['step'] == 4:  # 계산 완료 후
            btn_browse.config(bg=COLORS['completed'])
            btn_select_factor.config(bg=COLORS['completed'])
            if workflow_state['reverse_coding_done']:
                btn_reverse.config(bg=COLORS['completed'])
            btn_calculate.config(bg=COLORS['completed'])
            btn_prepare_next.config(bg=COLORS['active'], relief=tk.RAISED, state=tk.NORMAL)

        # 이전 애니메이션 정지 (메모리 누수 방지)
        if hasattr(update_button_states, 'glow_jobs'):
            for job in update_button_states.glow_jobs:
                try:
                    root.after_cancel(job)
                except:
                    pass
        update_button_states.glow_jobs = []

        # 반짝이는 효과를 위한 애니메이션 (다음 단계 버튼에만)
        def create_glow_effect(button, color1, color2, count=0):
            def glow():
                try:
                    if count >= 20:  # 최대 20번만 반복 (메모리 누수 방지)
                        button.config(bg=color1)  # 기본 색상으로 복원
                        return

                    current_bg = button.cget('bg')
                    next_bg = color2 if current_bg == color1 else color1
                    button.config(bg=next_bg)

                    job = root.after(800, lambda: create_glow_effect(button, color1, color2, count + 1))
                    update_button_states.glow_jobs.append(job)
                except:
                    pass
            glow()

        # 현재 단계의 다음 버튼에 반짝이는 효과
        if workflow_state['step'] == 1:
            create_glow_effect(btn_browse, COLORS['active'], COLORS['highlight'])
        elif workflow_state['step'] == 2:
            create_glow_effect(btn_select_factor, COLORS['active'], COLORS['highlight'])
        elif workflow_state['step'] == 3:
            create_glow_effect(btn_reverse, COLORS['glow'], COLORS['highlight'])
            create_glow_effect(btn_calculate, COLORS['glow'], COLORS['highlight'])
        elif workflow_state['step'] == 4:
            create_glow_effect(btn_prepare_next, COLORS['active'], COLORS['highlight'])

    except Exception as e:
        print(f"버튼 상태 업데이트 중 오류: {e}")


def suggest_factor_name(variables):
    """변수명들을 분석하여 요인명 추천"""
    if not variables:
        return "요인"

    # 공통 접두사/접미사 찾기
    common_parts = []

    # 1. 공통 접두사 찾기
    if len(variables) > 1:
        min_len = min(len(var) for var in variables)
        prefix = ""
        for i in range(min_len):
            if all(var[i] == variables[0][i] for var in variables):
                prefix += variables[0][i]
            else:
                break
        if len(prefix) >= 2:
            common_parts.append(prefix.rstrip('_0123456789'))

    # 2. 숫자와 '역' 제거하여 공통 부분 찾기
    cleaned_vars = []
    for var in variables:
        # 역코딩 변수 처리 ("역_" 접두사)
        cleaned = var.replace('역_', '')
        # 끝의 숫자(소수점 포함) 제거 (예: 문항1 → 문항, 문항1.1 → 문항)
        import re
        cleaned = re.sub(r'\d+\.?\d*$', '', cleaned)
        # 끝의 '_' 제거
        cleaned = cleaned.rstrip('_')
        cleaned_vars.append(cleaned)

    # 3. 가장 긴 공통 부분 찾기
    if cleaned_vars:
        # 가장 짧은 단어 기준으로 공통 부분 찾기
        shortest = min(cleaned_vars, key=len)
        for i in range(len(shortest), 0, -1):
            substring = shortest[:i]
            if all(substring in var for var in cleaned_vars):
                if len(substring) >= 2:
                    common_parts.append(substring)
                    break

    # 4. 추천 요인명 결정
    if common_parts:
        # 가장 긴 공통 부분 선택
        suggested_name = max(common_parts, key=len)
        # 불필요한 기호 제거
        suggested_name = suggested_name.replace('_', '').replace('-', '')
        return suggested_name if suggested_name else "요인"

    # 5. 공통 부분이 없으면 첫 번째 변수명 기반
    first_var = variables[0].replace('역_', '')
    import re
    base_name = re.sub(r'\d+\.?\d*$', '', first_var).rstrip('_')
    return base_name if len(base_name) >= 2 else "요인"


def find_similar_variables(target_var, all_variables):
    """클릭한 변수와 비슷한 이름의 변수들을 찾기"""
    import re

    # 제외할 변수 패턴들 (합계/평균만 제외, 역코딩 변수는 포함)
    exclude_patterns = [
        r'_합계$',        # 합계 변수
        r'_평균$',        # 평균 변수
        r'_mean$',        # 영문 평균
        r'_sum$',         # 영문 합계
        r'_total$'        # 영문 총계
    ]

    # 합계/평균 변수들만 제외
    filtered_vars = []
    for var in all_variables:
        is_excluded = False
        for pattern in exclude_patterns:
            if re.search(pattern, var):
                is_excluded = True
                break
        if not is_excluded:
            filtered_vars.append(var)

    # 클릭한 변수가 역코딩 변수인지 확인
    is_target_reverse = target_var.startswith('역_')

    # 클릭한 변수가 소수점을 포함하는지 확인
    target_clean = target_var.replace('역_', '')
    has_target_decimal = '.' in target_clean

    # 대상 변수에서 숫자(소수점 포함)와 역_ 제거하여 기본 패턴 추출
    target_base = target_var.replace('역_', '')  # 역_ 제거
    target_base = re.sub(r'\d+\.?\d*$', '', target_base).rstrip('_')  # 숫자(소수점 포함) 제거

    # 비슷한 변수들 찾기
    similar_vars = []
    for var in filtered_vars:
        var_is_reverse = var.startswith('역_')
        var_clean = var.replace('역_', '')
        has_var_decimal = '.' in var_clean

        var_base = var.replace('역_', '')  # 역_ 제거
        var_base = re.sub(r'\d+\.?\d*$', '', var_base).rstrip('_')  # 숫자(소수점 포함) 제거

        # 기본 패턴이 같고, 소수점 형식도 같은 변수들만
        if target_base == var_base and target_base and has_target_decimal == has_var_decimal:
            # 역코딩 변수 클릭 시 → 역코딩 변수들만 선택
            if is_target_reverse and var_is_reverse:
                similar_vars.append(var)
            # 원본 변수 클릭 시 → 원본 변수들만 선택 (역코딩 있어도 원본만)
            elif not is_target_reverse and not var_is_reverse:
                # 단, 같은 이름의 역코딩 변수가 존재하면 원본은 선택하지 않음
                reverse_version = f"역_{var}"
                if reverse_version not in all_variables:
                    similar_vars.append(var)

    return similar_vars

# 색상 테마 정의
COLORS = {
    'primary': '#87CEEB',      # 밝은 파란색 (메인)
    'secondary': '#DDA0DD',    # 밝은 보라색 (보조)
    'success': '#FFB347',      # 밝은 주황색 (성공)
    'warning': '#FF6B6B',      # 밝은 빨간색 (경고)
    'info': '#98D8C8',         # 밝은 청록색 (정보)
    'light': '#F5F5F5',        # 밝은 회색 (배경)
    'dark': '#2C3E50',         # 진한 회색 (텍스트)
    'white': '#FFFFFF',        # 흰색
    'button_text': '#000000',  # 버튼 텍스트 (검정색)

    # 단계별 강조 색상 추가
    'active': '#FFD700',       # 금색 (다음 단계 강조)
    'completed': '#90EE90',    # 연한 초록색 (완료된 단계)
    'disabled': '#D3D3D3',     # 회색 (비활성화)
    'highlight': '#FF4500',    # 주황색 (현재 활성 단계)
    'glow': '#32CD32'          # 라임 그린 (반짝이는 효과)
}

# 워크플로우 상태 관리
workflow_state = {
    'step': 1,  # 1: 파일선택, 2: 변수선택, 3: 역코딩, 4: 계산, 5: 완료
    'file_loaded': False,
    'variables_selected': False,
    'reverse_coding_done': False,
    'calculation_done': False
}


def select_file():
    """엑셀 파일 선택"""
    global df, original_column_order, original_file_path

    # GUI 컴포넌트 존재 확인
    if entry_file_path is None or listbox_factor_vars is None:
        messagebox.showerror("오류", "GUI가 초기화되지 않았습니다.")
        return

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if not file_path:
        return

    try:
        entry_file_path.delete(0, tk.END)
        entry_file_path.insert(0, file_path)

        # 원본 파일 경로 저장
        original_file_path = file_path

        # 파일 읽기 시도
        df = pd.read_excel(file_path)
        original_column_order = list(df.columns)

        # 데이터 유효성 검사
        if df.empty:
            messagebox.showwarning("경고", "파일이 비어있습니다.")
            return

        if len(df.columns) == 0:
            messagebox.showwarning("경고", "컬럼이 없는 파일입니다.")
            return

        # 리스트박스 업데이트
        listbox_factor_vars.delete(0, tk.END)

        for col in df.columns:
            listbox_factor_vars.insert(tk.END, col)

        messagebox.showinfo("파일 로드 완료", f"엑셀 파일이 성공적으로 로드되었습니다!\n행 수: {len(df)}, 열 수: {len(df.columns)}")
        update_result_text("✅ 파일 로드 완료\n📋 1단계: 같은 요인에 속하는 변수들을 선택하세요.")

        # 워크플로우 상태 업데이트
        workflow_state['step'] = 2
        workflow_state['file_loaded'] = True
        update_button_states()

    except FileNotFoundError:
        messagebox.showerror("오류", "파일을 찾을 수 없습니다.")
    except PermissionError:
        messagebox.showerror("오류", "파일에 접근할 권한이 없습니다.")
    except pd.errors.EmptyDataError:
        messagebox.showerror("오류", "빈 파일이거나 데이터가 없습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"파일을 열 수 없습니다: {str(e)}")


def select_factor_variables():
    """요인 변수 선택 및 자동으로 역코딩 창 열기"""
    global selected_factor_vars

    # 데이터 및 GUI 컴포넌트 확인
    if df is None:
        messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
        return

    if listbox_factor_vars is None:
        messagebox.showerror("오류", "GUI가 초기화되지 않았습니다.")
        return

    try:
        selected_indices = listbox_factor_vars.curselection()
        if not selected_indices:
            messagebox.showerror("오류", "같은 요인에 속하는 변수들을 선택하세요!")
            return

        # 선택된 변수들 가져오기 및 유효성 검사
        selected_factor_vars = []
        for idx in selected_indices:
            try:
                var_name = listbox_factor_vars.get(idx)
                if var_name and var_name in df.columns:
                    selected_factor_vars.append(var_name)
                else:
                    messagebox.showwarning("경고", f"변수 '{var_name}'이 데이터에 존재하지 않습니다.")
            except tk.TclError:
                messagebox.showerror("오류", "변수 선택 중 오류가 발생했습니다.")
                return

        if not selected_factor_vars:
            messagebox.showerror("오류", "유효한 변수가 선택되지 않았습니다.")
            return

        messagebox.showinfo("요인 변수 선택 완료",
                           f"선택된 변수: {', '.join(selected_factor_vars)}\n"
                           f"역코딩이 필요하면 '역코딩' 버튼을, 불필요하면 바로 '계산' 버튼을 클릭하세요.")

        # 워크플로우 상태 업데이트
        workflow_state['step'] = 3
        workflow_state['variables_selected'] = True
        update_button_states()

        # 선택 상태가 이미 화면에 실시간으로 표시되므로 별도 메시지 불필요
        # show_current_selection이 자동으로 계산 준비 상태까지 표시함

    except Exception as e:
        messagebox.showerror("오류", f"변수 선택 중 오류가 발생했습니다: {str(e)}")
        selected_factor_vars = []


def show_reverse_coding_dialog():
    """역코딩 변수 선택 팝업 창"""
    if not selected_factor_vars:
        messagebox.showerror("오류", "먼저 요인 변수들을 선택하세요!")
        return

    # 팝업 창 생성
    reverse_window = tk.Toplevel(root)
    reverse_window.title("역코딩 변수 선택")
    reverse_window.geometry("600x600")
    reverse_window.configure(bg=COLORS['light'])

    # 창을 항상 앞에 표시
    reverse_window.transient(root)
    reverse_window.grab_set()

    # 제목
    title_frame = tk.Frame(reverse_window, bg=COLORS['primary'], height=60)
    title_frame.pack(fill=tk.X, padx=10, pady=10)
    title_frame.pack_propagate(False)

    tk.Label(title_frame, text="🔄 역코딩할 변수 선택",
             font=("Arial", 14, "bold"), fg=COLORS['dark'],
             bg=COLORS['primary']).pack(expand=True)

    # 설명
    info_frame = tk.Frame(reverse_window, bg=COLORS['light'])
    info_frame.pack(fill=tk.X, padx=20, pady=10)

    tk.Label(info_frame, text="역코딩이 필요한 변수들을 선택하세요\n• 클릭: 개별 선택/해제 • 선택하지 않으면 원본 데이터 사용",
             font=("Arial", 10), fg=COLORS['dark'], bg=COLORS['light'],
             wraplength=450, justify=tk.LEFT).pack()

    # 변수 선택 리스트
    list_frame = tk.Frame(reverse_window, bg=COLORS['light'])
    list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    reverse_listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE,
                                font=("Arial", 14), height=10,
                                bg=COLORS['white'], fg=COLORS['dark'],
                                selectbackground=COLORS['info'],
                                activestyle='dotbox')
    reverse_listbox.pack(fill=tk.BOTH, expand=True)

    # 팝업창 클릭 선택 기능 (드래그 제거)
    def popup_on_click(event):
        widget = event.widget
        index = widget.nearest(event.y)

        if index < 0 or index >= widget.size():
            return "break"

        # 클릭: 토글 방식 (기존 선택 유지)
        current_selection = list(widget.curselection())
        if index in current_selection:
            widget.selection_clear(index)
        else:
            widget.selection_set(index)

        return "break"

    # 팝업창 기본 이벤트 차단
    def popup_disable_default(event):
        return "break"

    # 기본 이벤트들 모두 차단
    reverse_listbox.bind("<Button-1>", popup_disable_default)
    reverse_listbox.bind("<ButtonRelease-1>", popup_disable_default)
    reverse_listbox.bind("<B1-Motion>", popup_disable_default)
    reverse_listbox.bind("<Double-Button-1>", popup_disable_default)

    # 커스텀 클릭 이벤트만 허용
    reverse_listbox.bind("<Button-1>", popup_on_click)

    for var in selected_factor_vars:
        reverse_listbox.insert(tk.END, var)


    # 역코딩 설정 프레임
    settings_frame = tk.Frame(reverse_window, bg=COLORS['light'])
    settings_frame.pack(fill=tk.X, padx=20, pady=10)

    tk.Label(settings_frame, text="역코딩 설정",
             font=("Arial", 11, "bold"), fg=COLORS['dark'],
             bg=COLORS['light']).pack(anchor=tk.W)

    settings_inner = tk.Frame(settings_frame, bg=COLORS['light'])
    settings_inner.pack(fill=tk.X, pady=5)

    tk.Label(settings_inner, text="최대값:", font=("Arial", 10),
             fg=COLORS['dark'], bg=COLORS['light']).pack(side=tk.LEFT)
    max_entry = tk.Entry(settings_inner, width=8, font=("Arial", 10))
    max_entry.insert(0, "5")
    max_entry.pack(side=tk.LEFT, padx=(5, 15))

    tk.Label(settings_inner, text="최소값:", font=("Arial", 10),
             fg=COLORS['dark'], bg=COLORS['light']).pack(side=tk.LEFT)
    min_entry = tk.Entry(settings_inner, width=8, font=("Arial", 10))
    min_entry.insert(0, "1")
    min_entry.pack(side=tk.LEFT, padx=5)

    # 버튼 프레임
    button_frame = tk.Frame(reverse_window, bg=COLORS['light'])
    button_frame.pack(fill=tk.X, padx=20, pady=20)

    def apply_reverse_coding():
        selected_indices = reverse_listbox.curselection()
        reverse_vars = [reverse_listbox.get(idx) for idx in selected_indices]

        if not reverse_vars:
            messagebox.showinfo("정보", "역코딩할 변수를 선택해주세요!")
            return

        try:
            max_value = float(max_entry.get())
            min_value = float(min_entry.get())
        except ValueError:
            messagebox.showerror("오류", "최대값과 최소값을 올바르게 입력하세요!")
            return

        # 역코딩 시작 메시지 즉시 표시
        update_result_text(f"🚀 역코딩을 시작합니다...\n📝 대상 변수: {', '.join(reverse_vars)}\n⚙️ 최대값: {max_value}, 최소값: {min_value}")

        # GUI 즉시 업데이트
        root.update_idletasks()

        perform_reverse_coding_internal(reverse_vars, max_value, min_value)
        reverse_window.destroy()

    tk.Button(button_frame, text="역코딩 실행", command=apply_reverse_coding,
              bg=COLORS['success'], fg=COLORS['button_text'], font=("Arial", 11, "bold"),
              padx=20, pady=8).pack(side=tk.RIGHT, padx=5)

    tk.Button(button_frame, text="취소", command=reverse_window.destroy,
              bg=COLORS['warning'], fg=COLORS['button_text'], font=("Arial", 11, "bold"),
              padx=20, pady=8).pack(side=tk.RIGHT)


def perform_reverse_coding_internal(reverse_vars, max_value, min_value):
    """내부 역코딩 수행 함수"""
    global df, reverse_coded_columns

    # 입력 데이터 검증
    if not reverse_vars:
        messagebox.showinfo("정보", "역코딩할 변수가 선택되지 않았습니다.")
        update_result_text("ℹ️ 역코딩할 변수가 없습니다. 원본 데이터로 합계/평균을 계산합니다.")
        return

    if df is None:
        messagebox.showerror("오류", "데이터가 로드되지 않았습니다.")
        return

    # 숫자 값 검증
    try:
        max_value = float(max_value)
        min_value = float(min_value)
        if max_value <= min_value:
            messagebox.showerror("오류", "최대값은 최소값보다 커야 합니다.")
            return
    except (ValueError, TypeError):
        messagebox.showerror("오류", "최대값과 최소값은 숫자여야 합니다.")
        return

    try:
        # 역코딩할 변수들이 실제 데이터에 존재하는지 먼저 확인
        missing_vars = [var for var in reverse_vars if var not in df.columns]
        if missing_vars:
            messagebox.showerror("오류", f"다음 변수들이 데이터에 존재하지 않습니다: {', '.join(missing_vars)}")
            return

        # 역코딩 수행 (진행상황 실시간 표시)
        completed_vars = []
        failed_vars = []

        for i, var in enumerate(reverse_vars):
            try:
                if var in df.columns:
                    reverse_col_name = f"역_{var}"

                    # 데이터 타입 확인 및 변환
                    if not pd.api.types.is_numeric_dtype(df[var]):
                        # 숫자로 변환 시도
                        df[var] = pd.to_numeric(df[var], errors='coerce')

                    # 결측치 처리 확인
                    if df[var].isna().any():
                        messagebox.showwarning("경고", f"변수 '{var}'에 결측치가 있습니다. 결측치는 그대로 유지됩니다.")

                    # 역코딩 공식: 최대값 + 최소값 - 원본값
                    df[reverse_col_name] = max_value + min_value - df[var]
                    reverse_coded_columns[var] = reverse_col_name
                    completed_vars.append(var)

                    # 진행상황 실시간 업데이트
                    progress_text = f"🔄 역코딩 진행중...\n📊 진행률: {i+1}/{len(reverse_vars)}\n✅ 완료된 변수: {', '.join(completed_vars)}"
                    if failed_vars:
                        progress_text += f"\n❌ 실패한 변수: {', '.join(failed_vars)}"
                    update_result_text(progress_text)

                    # 중간 과정에서도 메인 리스트 업데이트하여 새로운 역코딩 변수가 즉시 보이게 함
                    if root is not None:
                        refresh_main_variable_list()
                        root.update_idletasks()

                    # 짧은 딜레이로 사용자가 진행과정을 볼 수 있게 함
                    import time
                    time.sleep(0.1)

                else:
                    failed_vars.append(var)
                    messagebox.showerror("오류", f"변수 '{var}'를 데이터에서 찾을 수 없습니다.")

            except Exception as var_error:
                failed_vars.append(var)
                print(f"변수 '{var}' 역코딩 중 오류: {str(var_error)}")
                continue

        # 컬럼 순서 재배치 (역코딩 변수를 원본 바로 뒤에)
        new_columns = []
        for col in original_column_order:
            new_columns.append(col)
            if col in reverse_coded_columns:
                new_columns.append(reverse_coded_columns[col])

        # 새로 생긴 합계/평균 컬럼들도 추가
        for col in df.columns:
            if col not in new_columns:
                new_columns.append(col)

        # 컬럼 순서 재배치 (전역 변수 명시적 업데이트)
        df = df[new_columns].copy()

        messagebox.showinfo("역코딩 완료",
                           f"✅ 역코딩 완료!\n📊 역코딩된 변수: {', '.join(reverse_vars)}\n"
                           f"🔍 새로 생성된 변수: {', '.join([f'역_{var}' for var in reverse_vars])}")

        # 역코딩 완료 후 선택 상태 업데이트: 역코딩된 변수들을 선택하고 원본은 해제
        updated_selected_vars = []
        for var in selected_factor_vars:
            if var in reverse_coded_columns:
                # 역코딩된 변수가 있으면 역코딩 변수를 선택 목록에 추가
                updated_selected_vars.append(reverse_coded_columns[var])
            else:
                # 역코딩되지 않은 변수는 원본 그대로 유지
                updated_selected_vars.append(var)

        # 메인 화면의 변수 리스트 업데이트하면서 역코딩 변수들을 선택 상태로 설정
        root.update_idletasks()  # GUI 즉시 업데이트
        refresh_main_variable_list_with_selection(updated_selected_vars)
        root.update_idletasks()  # 한번 더 업데이트로 확실히

        # 워크플로우 상태 업데이트
        workflow_state['reverse_coding_done'] = True
        update_button_states()

        # 선택 상태가 업데이트된 후 계산 준비 상태도 자동으로 표시됨 (show_current_selection에서 처리)

    except Exception as e:
        messagebox.showerror("오류", f"역코딩 중 오류가 발생했습니다:\n{e}")
        update_result_text(f"❌ 역코딩 실패\n오류: {str(e)}\n다시 시도하거나 관리자에게 문의하세요.")


def calculate_factor_statistics():
    """요인 합계 및 평균 계산"""
    global df, selected_factor_vars, reverse_coded_columns

    # 기본 데이터 검증
    if df is None:
        messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
        return

    if df.empty:
        messagebox.showerror("오류", "데이터가 비어있습니다!")
        return

    if listbox_factor_vars is None:
        messagebox.showerror("오류", "GUI가 초기화되지 않았습니다.")
        return

    # 현재 선택 상태 저장 (계산 중에도 유지하기 위해)
    current_selected_vars = []
    for idx in listbox_factor_vars.curselection():
        current_selected_vars.append(listbox_factor_vars.get(idx))

    # 실제 계산에 사용할 변수들 = 현재 왼쪽에서 선택된 변수들
    if not current_selected_vars:
        messagebox.showerror("오류", "먼저 요인 변수들을 선택하세요!")
        return

    # 요인명 자동 추천 (실제 선택된 변수들 기반으로)
    # 역코딩 변수명에서 '역_' 제거하여 원본 이름으로 추천
    original_names = []
    for var in current_selected_vars:
        if var.startswith('역_'):
            original_names.append(var[2:])  # '역_' 제거
        else:
            original_names.append(var)

    suggested_name = suggest_factor_name(original_names)

    # 요인명 입력받기 (추천명을 기본값으로)
    factor_name = simpledialog.askstring("요인명 입력",
                                        f"합계/평균 변수에 사용할 요인명을 입력하세요:\n(추천: {suggested_name})",
                                        initialvalue=suggested_name)
    if not factor_name:
        return

    # 계산에 사용할 변수 리스트 = 현재 선택된 변수들
    calculation_vars = current_selected_vars
    used_vars_info = []
    detailed_info = []

    # 선택된 변수들이 실제 데이터에 존재하는지 검증
    missing_vars = []
    invalid_vars = []

    for var in current_selected_vars:
        if var not in df.columns:
            missing_vars.append(var)
            continue

        # 데이터 타입 검증
        try:
            if not pd.api.types.is_numeric_dtype(df[var]):
                # 숫자로 변환 시도
                numeric_data = pd.to_numeric(df[var], errors='coerce')
                if numeric_data.isna().all():
                    invalid_vars.append(var)
                    continue
        except Exception:
            invalid_vars.append(var)
            continue

        # 유효한 변수 처리
        if var.startswith('역_'):
            # 역코딩 변수
            original_name = var[2:]  # '역_' 제거
            used_vars_info.append(f"📊 {original_name} → {var} (역코딩 데이터)")
            detailed_info.append(f"{original_name}(역코딩)")
        else:
            # 원본 변수
            used_vars_info.append(f"📊 {var} (원본 데이터)")
            detailed_info.append(f"{var}(원본)")

    # 오류가 있으면 사용자에게 알림
    if missing_vars:
        messagebox.showerror("오류", f"다음 변수들이 데이터에 존재하지 않습니다: {', '.join(missing_vars)}")
        return

    if invalid_vars:
        messagebox.showerror("오류", f"다음 변수들은 숫자 데이터가 아닙니다: {', '.join(invalid_vars)}")
        return

    # 유효한 변수가 남아있는지 확인
    valid_vars = [var for var in current_selected_vars if var not in missing_vars and var not in invalid_vars]
    if len(valid_vars) < 2:
        messagebox.showwarning("경고", "계산을 위해서는 최소 2개 이상의 유효한 변수가 필요합니다.")
        return

    calculation_vars = valid_vars

    # 계산 전 미리보기 표시 (간결하게) - 선택 상태 유지
    preview_text = f"""
📊 {factor_name} 합계/평균 계산 미리보기

생성될 변수:
  • {factor_name}_합계
  • {factor_name}_평균

계산 공식:
  • 합계 = {' + '.join(detailed_info)}
  • 평균 = ({' + '.join(detailed_info)}) ÷ {len(calculation_vars)}

사용할 데이터:
""" + "\n".join(used_vars_info) + f"""

총 {len(calculation_vars)}개 변수로 계산됩니다.
"""

    update_result_text(preview_text)
    root.update_idletasks()

    # 1초 대기 후 실제 계산 시작
    import time
    time.sleep(1.0)

    # 계산 진행 메시지 - 선택 상태 유지
    update_result_text(f"🔄 계산 실행 중...\n⏳ {factor_name} 합계 및 평균을 계산하고 있습니다...")
    root.update_idletasks()

    # 합계 및 평균 계산
    sum_col_name = f"{factor_name}_합계"
    mean_col_name = f"{factor_name}_평균"

    df[sum_col_name] = df[calculation_vars].sum(axis=1)
    df[mean_col_name] = df[calculation_vars].mean(axis=1)

    # 간결한 최종 결과 메시지 - 실제 사용된 변수들 표시
    result_message = f"""
✅ {factor_name} 계산 완료!

계산에 사용된 변수들:
""" + '\n'.join([f"  [{i+1}] {var}" for i, var in enumerate(current_selected_vars)]) + f"""

현황: 원본 {len([v for v in current_selected_vars if not v.startswith('역_')])}개, 역코딩 {len([v for v in current_selected_vars if v.startswith('역_')])}개, 총 {len(current_selected_vars)}개

다음 단계: "다음 요인 계산 준비" 버튼 클릭
"""

    messagebox.showinfo("계산 완료", f"✅ {factor_name} 합계 및 평균 계산 완료!\n생성된 변수: {sum_col_name}, {mean_col_name}\n\n🔄 자동으로 1단계로 돌아갑니다.")
    update_result_text(result_message)

    # 자동으로 다음 요인 계산 준비 (1단계로 돌아가기)
    auto_prepare_next_factor()

    # 워크플로우 상태 업데이트 (1단계로 초기화)
    workflow_state['step'] = 1
    workflow_state['variables_selected'] = False
    workflow_state['reverse_coding_done'] = False
    workflow_state['calculation_done'] = False
    update_button_states()

    # 메인 변수 리스트 업데이트 (선택 해제)
    refresh_main_variable_list()
    root.update_idletasks()




def save_to_excel():
    """결과를 엑셀 파일로 저장"""
    global original_file_path

    if df is None:
        messagebox.showerror("오류", "먼저 분석을 실행하세요!")
        return

    # 기본 파일명 생성 (원본 파일명 + "_변수 계산 완료")
    default_filename = ""
    if original_file_path:
        import os
        # 원본 파일의 디렉토리와 파일명 분리
        file_dir = os.path.dirname(original_file_path)
        file_name = os.path.basename(original_file_path)

        # 확장자 분리
        name_without_ext, ext = os.path.splitext(file_name)

        # 새 파일명 생성: 원본이름_변수 계산 완료.xlsx
        new_filename = f"{name_without_ext}_변수 계산 완료.xlsx"
        default_filename = os.path.join(file_dir, new_filename)

    # 저장 대화상자 (기본 파일명 포함)
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile=os.path.basename(default_filename) if default_filename else "변수 계산 완료.xlsx",
        initialdir=os.path.dirname(default_filename) if default_filename else None
    )

    if not save_path:
        return

    try:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("저장 완료", f"✅ 결과가 저장되었습니다!\n📁 {save_path}")
    except Exception as e:
        messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{e}")


def prepare_next_factor():
    """다음 요인 계산 준비 - 선택 초기화하되 데이터는 유지"""
    global selected_factor_vars

    if df is None:
        messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
        return

    # 선택 초기화
    selected_factor_vars = []
    listbox_factor_vars.selection_clear(0, tk.END)

    # 기본 안내 메시지로 복원 (선택이 없으면 show_current_selection에서 기본 메시지 표시)
    update_result_text("✅ 다음 요인 계산 준비 완료!\n\n📋 1단계: 새로운 요인에 속하는 변수들을 선택하세요\n\n🎯 이전에 계산된 합계/평균 변수들과 역코딩 변수들은 그대로 유지됩니다.")

    # 워크플로우 상태 리셋 (새로운 요인 시작)
    workflow_state['step'] = 2  # 변수 선택 단계로
    workflow_state['variables_selected'] = False
    workflow_state['calculation_done'] = False
    # reverse_coding_done과 file_loaded는 유지
    update_button_states()

    messagebox.showinfo("준비 완료", "✅ 다음 요인 계산 준비가 완료되었습니다!\n\n새로운 요인에 속하는 변수들을 선택해주세요.")


def auto_prepare_next_factor():
    """자동으로 다음 요인 계산 준비 - 메시지 박스 없이 실행"""
    global selected_factor_vars

    if df is None:
        return

    # 선택 초기화
    selected_factor_vars = []
    listbox_factor_vars.selection_clear(0, tk.END)

    # 1단계 안내 메시지
    update_result_text("🔄 자동으로 1단계로 돌아갔습니다!\n\n📋 1단계: 새로운 요인에 속하는 변수들을 선택하세요\n\n🎯 이전에 계산된 합계/평균 변수들과 역코딩 변수들은 그대로 유지됩니다.\n\n✨ 다음 요인 변수들을 선택해서 계속 진행하세요.")


def analyze_variable_range(var_name):
    """변수의 데이터를 분석하여 최대/최소값 추정"""
    if df is None or var_name not in df.columns:
        return 5, 1  # 기본값

    try:
        # 숫자형 데이터로 변환 시도
        data = pd.to_numeric(df[var_name], errors='coerce').dropna()

        if len(data) == 0:
            return 5, 1  # 데이터 없으면 기본값

        actual_min = int(data.min())
        actual_max = int(data.max())

        # 일반적인 척도 범위로 보정
        if actual_max <= 5 and actual_min >= 1:
            return 5, 1  # 1-5 척도
        elif actual_max <= 7 and actual_min >= 1:
            return 7, 1  # 1-7 척도
        elif actual_max <= 10 and actual_min >= 1:
            return 10, 1  # 1-10 척도
        else:
            return actual_max, actual_min  # 실제 범위 사용

    except:
        return 5, 1  # 오류 시 기본값


def auto_group_variables():
    """변수들을 자동으로 그룹핑하여 요인별로 분류"""
    if df is None:
        return {}

    import re

    # 모든 변수명에서 패턴 추출
    variable_groups = {}

    # 합계/평균 변수 제외
    exclude_patterns = [r'_합계$', r'_평균$', r'_mean$', r'_sum$', r'_total$']
    filtered_vars = []

    for var in df.columns:
        is_excluded = False
        for pattern in exclude_patterns:
            if re.search(pattern, str(var)):
                is_excluded = True
                break
        if not is_excluded:
            filtered_vars.append(var)

    # 변수들을 패턴별로 그룹핑
    for var in filtered_vars:
        var_str = str(var)

        # 역코딩 변수 처리 ("역_" 접두사)
        clean_var = var_str.replace('역_', '')

        # 숫자(소수점 포함) 제거하여 기본 패턴 추출
        base_pattern = re.sub(r'\d+\.?\d*$', '', clean_var).rstrip('_')

        if len(base_pattern) >= 2:  # 의미있는 패턴만
            if base_pattern not in variable_groups:
                variable_groups[base_pattern] = []
            variable_groups[base_pattern].append(var)

    # 2개 이상의 변수가 있는 그룹만 반환
    return {k: v for k, v in variable_groups.items() if len(v) >= 2}


def quick_calculation():
    """빠른 계산 - 트리뷰와 분할 화면으로 새롭게 구현"""
    global df

    if df is None:
        messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
        return

    # ttk 스타일 적용을 위한 import
    from tkinter import ttk

    # 자동 그룹핑
    groups = auto_group_variables()

    if not groups:
        messagebox.showwarning("경고", "자동으로 그룹핑할 수 있는 변수가 없습니다.\n수동으로 변수를 선택해서 계산해주세요.")
        return

    # 전체 화면 크기 설정
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = min(1400, screen_width - 50)
    window_height = min(900, screen_height - 50)

    # 새로운 창 생성
    calc_window = tk.Toplevel(root)
    calc_window.title("⚡ 빠른 계산 - 스마트 선택")
    calc_window.geometry(f"{window_width}x{window_height}")
    calc_window.configure(bg=COLORS['light'])
    calc_window.transient(root)
    calc_window.grab_set()

    # 중앙 정렬
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    calc_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # 제목 영역
    title_frame = tk.Frame(calc_window, bg=COLORS['highlight'], height=70)
    title_frame.pack(fill=tk.X, padx=10, pady=10)
    title_frame.pack_propagate(False)

    tk.Label(title_frame, text="⚡ 빠른 계산 - 스마트 선택",
             font=("Arial", 20, "bold"), fg=COLORS['button_text'],
             bg=COLORS['highlight']).pack(expand=True)

    # 메인 분할 영역 (PanedWindow 사용)
    main_paned = tk.PanedWindow(calc_window, orient=tk.HORIZONTAL,
                               sashwidth=5, sashrelief=tk.RAISED, bg=COLORS['dark'])
    main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

    # 왼쪽 영역: 카테고리 트리뷰
    left_frame = tk.Frame(main_paned, bg=COLORS['white'], relief=tk.RAISED, bd=2)
    left_frame.pack(fill=tk.BOTH, expand=True)

    # 왼쪽 제목
    left_title = tk.Label(left_frame, text="📂 변수 그룹 선택",
                         font=("Arial", 16, "bold"), fg=COLORS['dark'],
                         bg=COLORS['info'], pady=10)
    left_title.pack(fill=tk.X)

    # 트리뷰 생성
    tree_frame = tk.Frame(left_frame, bg=COLORS['white'])
    tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 트리뷰와 스크롤바
    tree = ttk.Treeview(tree_frame, height=20)
    tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll.set)

    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # 오른쪽 영역: 선택된 항목 설정
    right_frame = tk.Frame(main_paned, bg=COLORS['white'], relief=tk.RAISED, bd=2)
    right_frame.pack(fill=tk.BOTH, expand=True)

    # 오른쪽 제목
    right_title = tk.Label(right_frame, text="⚙️ 선택된 그룹 설정",
                          font=("Arial", 16, "bold"), fg=COLORS['dark'],
                          bg=COLORS['success'], pady=10)
    right_title.pack(fill=tk.X)

    # 오른쪽 스크롤 영역
    right_canvas = tk.Canvas(right_frame, bg=COLORS['white'], highlightthickness=0)
    right_scrollbar = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=right_canvas.yview)
    right_scrollable = tk.Frame(right_canvas, bg=COLORS['white'])

    right_scrollable.bind("<Configure>", lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all")))
    right_canvas.create_window((0, 0), window=right_scrollable, anchor="nw")
    right_canvas.configure(yscrollcommand=right_scrollbar.set)

    right_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
    right_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # 분할 창에 추가
    main_paned.add(left_frame, minsize=400)
    main_paned.add(right_frame, minsize=500)

    # 데이터 저장
    selected_groups = {}  # 선택된 그룹들의 설정
    group_frames = {}     # 오른쪽에 생성된 프레임들

    # 트리뷰에 그룹과 변수 추가
    for group_name, variables in groups.items():
        # 그룹 아이템 추가
        group_item = tree.insert("", "end", text=f"📊 {group_name} ({len(variables)}개)",
                               values=(group_name, "group"), open=False)

        # 변수들을 하위 아이템으로 추가
        for var in variables:
            tree.insert(group_item, "end", text=f"📋 {var}",
                       values=(var, "variable", group_name))

    def on_tree_select(event):
        """트리뷰 선택 이벤트"""
        selection = tree.selection()
        if not selection:
            return

        item = selection[0]
        values = tree.item(item, "values")

        if len(values) >= 2 and values[1] == "group":
            group_name = values[0]
            add_group_to_right(group_name, groups[group_name])

    def add_group_to_right(group_name, variables):
        """오른쪽에 그룹 설정 추가"""
        if group_name in selected_groups:
            return  # 이미 추가된 그룹

        # 그룹별 데이터 분석
        sample_var = variables[0]
        group_max, group_min = analyze_variable_range(sample_var)

        # 그룹 프레임 생성
        group_frame = tk.Frame(right_scrollable, bg=COLORS['info'], relief=tk.RAISED, bd=3)
        group_frame.pack(fill=tk.X, padx=5, pady=5)

        # 그룹 헤더
        header_frame = tk.Frame(group_frame, bg=COLORS['primary'])
        header_frame.pack(fill=tk.X, padx=5, pady=5)

        # 그룹명과 제거 버튼
        header_left = tk.Frame(header_frame, bg=COLORS['primary'])
        header_left.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(header_left, text=f"📊 {group_name}",
                font=("Arial", 14, "bold"), fg=COLORS['dark'], bg=COLORS['primary']).pack(side=tk.LEFT)

        def remove_group():
            group_frame.destroy()
            del selected_groups[group_name]
            del group_frames[group_name]

        tk.Button(header_frame, text="❌", command=remove_group,
                 bg=COLORS['warning'], fg=COLORS['button_text'],
                 font=("Arial", 12, "bold"), padx=10).pack(side=tk.RIGHT)

        # 역코딩 설정
        range_frame = tk.Frame(header_frame, bg=COLORS['primary'])
        range_frame.pack(side=tk.RIGHT, padx=10)

        tk.Label(range_frame, text="역코딩 범위:", font=("Arial", 11, "bold"),
                fg=COLORS['dark'], bg=COLORS['primary']).pack(side=tk.LEFT)

        max_var = tk.StringVar(value=str(group_max))
        tk.Entry(range_frame, textvariable=max_var, width=5,
                font=("Arial", 12, "bold"), justify=tk.CENTER).pack(side=tk.LEFT, padx=2)

        tk.Label(range_frame, text="~", font=("Arial", 11),
                fg=COLORS['dark'], bg=COLORS['primary']).pack(side=tk.LEFT)

        min_var = tk.StringVar(value=str(group_min))
        tk.Entry(range_frame, textvariable=min_var, width=5,
                font=("Arial", 12, "bold"), justify=tk.CENTER).pack(side=tk.LEFT, padx=2)

        # 변수 목록
        vars_frame = tk.Frame(group_frame, bg=COLORS['white'])
        vars_frame.pack(fill=tk.X, padx=5, pady=5)

        var_checkboxes = {}
        reverse_checkboxes = {}

        # 변수들을 3열로 배치
        for k, var in enumerate(variables):
            row_num = k // 3
            col_num = k % 3

            if col_num == 0:
                var_row = tk.Frame(vars_frame, bg=COLORS['white'])
                var_row.pack(fill=tk.X, pady=2)

            var_container = tk.Frame(var_row, bg=COLORS['light'], relief=tk.GROOVE, bd=1)
            var_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)

            # 변수 포함 체크박스
            include_var = tk.BooleanVar(value=True)
            include_check = tk.Checkbutton(var_container, text=f"📋 {var}",
                                         font=("Arial", 11, "bold"), fg=COLORS['dark'],
                                         bg=COLORS['light'], variable=include_var)
            include_check.pack(anchor=tk.W, padx=3, pady=2)

            # 역코딩 체크박스
            reverse_var = tk.BooleanVar(value=False)
            reverse_check = tk.Checkbutton(var_container, text="🔄 역코딩",
                                         font=("Arial", 10), fg=COLORS['dark'],
                                         bg=COLORS['light'], variable=reverse_var)
            reverse_check.pack(anchor=tk.W, padx=3, pady=1)

            var_checkboxes[var] = include_var
            reverse_checkboxes[var] = reverse_var

        # 그룹 데이터 저장
        selected_groups[group_name] = {
            'max_val': max_var,
            'min_val': min_var,
            'var_included': var_checkboxes,
            'var_reverse': reverse_checkboxes,
            'variables': variables
        }
        group_frames[group_name] = group_frame

        # 캔버스 스크롤 영역 업데이트
        right_canvas.update_idletasks()
        right_canvas.configure(scrollregion=right_canvas.bbox("all"))

    # 트리뷰 이벤트 바인딩
    tree.bind("<<TreeviewSelect>>", on_tree_select)

    # 하단 버튼 영역
    bottom_frame = tk.Frame(calc_window, bg=COLORS['secondary'], height=80)
    bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
    bottom_frame.pack_propagate(False)

    button_container = tk.Frame(bottom_frame, bg=COLORS['secondary'])
    button_container.pack(expand=True)

    def execute_calculation():
        """계산 실행"""
        if not selected_groups:
            messagebox.showwarning("경고", "선택된 그룹이 없습니다!")
            return

        try:
            # 설정 수집
            calculation_data = {}
            for group_name, data in selected_groups.items():
                included_vars = {}
                for var in data['variables']:
                    if data['var_included'][var].get():
                        included_vars[var] = {
                            'reverse': data['var_reverse'][var].get()
                        }

                if included_vars:
                    calculation_data[group_name] = {
                        'variables': included_vars,
                        'max_val': float(data['max_val'].get()),
                        'min_val': float(data['min_val'].get())
                    }

            if not calculation_data:
                messagebox.showwarning("경고", "포함될 변수가 없습니다!")
                return

            calc_window.destroy()
            perform_final_bulk_calculation(calculation_data)

        except ValueError:
            messagebox.showerror("오류", "역코딩 값은 숫자여야 합니다.")
        except Exception as e:
            messagebox.showerror("오류", f"계산 실행 오류: {str(e)}")

    # 버튼들
    tk.Button(button_container, text="🚀 계산 실행", command=execute_calculation,
              bg=COLORS['highlight'], fg=COLORS['button_text'], font=("Arial", 18, "bold"),
              padx=50, pady=20, relief=tk.RAISED, bd=4).pack(side=tk.LEFT, padx=20)

    tk.Button(button_container, text="❌ 취소", command=calc_window.destroy,
              bg=COLORS['warning'], fg=COLORS['button_text'], font=("Arial", 16, "bold"),
              padx=40, pady=15, relief=tk.RAISED, bd=3).pack(side=tk.LEFT, padx=20)

    # 부드러운 마우스 휠 스크롤 지원 (개선된 버전)
    def smooth_mousewheel(event):
        """부드러운 마우스 휠 스크롤 핸들러"""
        try:
            # 마우스가 오른쪽 캔버스 영역에 있는지 확인
            widget_under_mouse = event.widget.winfo_containing(event.x_root, event.y_root)

            # 오른쪽 캔버스나 그 자식 위젯들에서만 스크롤 적용
            if (widget_under_mouse == right_canvas or
                widget_under_mouse == right_scrollable or
                str(widget_under_mouse).startswith(str(right_scrollable))):

                # 스크롤 방향과 양 계산 (부드럽게)
                if hasattr(event, 'delta'):
                    # 맥북 트랙패드: delta 값을 더 세밀하게 처리
                    if event.delta != 0:
                        scroll_amount = max(1, min(3, abs(event.delta) // 40))
                        scroll_direction = -1 if event.delta > 0 else 1

                        # 부드러운 스크롤을 위해 작은 단위로 여러 번
                        for _ in range(scroll_amount):
                            right_canvas.yview_scroll(scroll_direction, "units")
                            right_canvas.update_idletasks()

                elif hasattr(event, 'num'):
                    # Linux 마우스 휠
                    if event.num == 4:
                        for _ in range(3):
                            right_canvas.yview_scroll(-1, "units")
                    elif event.num == 5:
                        for _ in range(3):
                            right_canvas.yview_scroll(1, "units")
                else:
                    # 기본 처리 (Windows)
                    scroll_units = int(-1 * (event.delta / 120))
                    right_canvas.yview_scroll(scroll_units, "units")

        except Exception as e:
            print(f"스크롤 처리 오류: {e}")

        return "break"

    # 스크롤 이벤트 바인딩 (다중 이벤트 지원)
    def setup_scroll_bindings():
        """스크롤 바인딩 설정"""
        try:
            # 다양한 스크롤 이벤트
            scroll_events = [
                "<MouseWheel>", "<Button-4>", "<Button-5>",
                "<Shift-MouseWheel>", "<Control-MouseWheel>"
            ]

            # 여러 위젯에 스크롤 바인딩
            for event in scroll_events:
                calc_window.bind_all(event, smooth_mousewheel)
                right_canvas.bind(event, smooth_mousewheel)
                right_scrollable.bind(event, smooth_mousewheel)

            # 캔버스에 포커스 설정
            right_canvas.focus_set()

            print("부드러운 스크롤 바인딩 완료")

        except Exception as e:
            print(f"스크롤 바인딩 오류: {e}")

    # 키보드 스크롤 지원
    def on_key_scroll(event):
        """키보드 스크롤 핸들러"""
        try:
            if event.keysym == "Up":
                right_canvas.yview_scroll(-3, "units")
            elif event.keysym == "Down":
                right_canvas.yview_scroll(3, "units")
            elif event.keysym == "Page_Up":
                right_canvas.yview_scroll(-10, "units")
            elif event.keysym == "Page_Down":
                right_canvas.yview_scroll(10, "units")
            elif event.keysym == "Home":
                right_canvas.yview_moveto(0)
            elif event.keysym == "End":
                right_canvas.yview_moveto(1)
        except Exception as e:
            print(f"키보드 스크롤 오류: {e}")

    # 키보드 바인딩
    def setup_keyboard_bindings():
        """키보드 스크롤 바인딩"""
        try:
            keyboard_events = [
                "<Up>", "<Down>", "<Page_Up>", "<Page_Down>",
                "<Home>", "<End>"
            ]

            for event in keyboard_events:
                calc_window.bind_all(event, on_key_scroll)

            print("키보드 스크롤 바인딩 완료")

        except Exception as e:
            print(f"키보드 바인딩 오류: {e}")

    # 창 완전 생성 후 스크롤 설정
    calc_window.after(100, setup_scroll_bindings)
    calc_window.after(150, setup_keyboard_bindings)


def perform_final_bulk_calculation(groups_data):
    """최종 일괄 계산 - 사용자 설정 반영"""
    global df, reverse_coded_columns

    try:
        total_groups = len(groups_data)
        completed_groups = 0
        total_created_vars = 0

        update_result_text(f"⚡ 고급 계산 시작!\n📊 총 {total_groups}개 그룹 처리 중\n")
        root.update_idletasks()

        detailed_results = []

        for group_name, group_settings in groups_data.items():
            # 진행상황 표시
            update_result_text(f"🔄 고급 계산 진행중...\n📊 진행률: {completed_groups+1}/{total_groups}\n🎯 현재 처리: {group_name}")
            root.update_idletasks()

            group_max = group_settings['max_val']
            group_min = group_settings['min_val']
            var_settings = group_settings['variables']

            # 1. 사용자 설정에 따른 역코딩
            reverse_vars = []
            included_vars = []

            # 원본 변수 바로 뒤에 역코딩 변수 배치를 위한 순서 처리
            current_columns = list(df.columns)

            for var_name, settings in var_settings.items():
                if settings['reverse']:  # 역코딩 선택된 변수
                    reverse_col_name = f"역_{var_name}"  # 빠른 계산에서는 "역_" 접두사 사용
                    if var_name in df.columns and reverse_col_name not in df.columns:
                        try:
                            if not pd.api.types.is_numeric_dtype(df[var_name]):
                                df[var_name] = pd.to_numeric(df[var_name], errors='coerce')

                            # 역코딩 계산
                            df[reverse_col_name] = group_max + group_min - df[var_name]
                            reverse_coded_columns[var_name] = reverse_col_name
                            reverse_vars.append(var_name)
                            included_vars.append(reverse_col_name)

                            # 원본 변수 바로 뒤에 역코딩 변수 배치
                            if var_name in current_columns:
                                var_index = current_columns.index(var_name)
                                current_columns.insert(var_index + 1, reverse_col_name)
                            else:
                                current_columns.append(reverse_col_name)

                        except Exception as e:
                            print(f"역코딩 오류 {var_name}: {e}")
                            included_vars.append(var_name)
                else:
                    included_vars.append(var_name)

            # 컬럼 순서 재배치 (역코딩 변수가 원본 바로 뒤에 오도록)
            df = df.reindex(columns=current_columns)

            # 2. 합계 및 평균 계산
            if len(included_vars) >= 1:  # 1개 변수라도 계산 허용
                sum_col_name = f"{group_name}_합계"
                mean_col_name = f"{group_name}_평균"

                # 중복 체크
                counter = 1
                while sum_col_name in df.columns:
                    sum_col_name = f"{group_name}{counter}_합계"
                    mean_col_name = f"{group_name}{counter}_평균"
                    counter += 1

                try:
                    df[sum_col_name] = df[included_vars].sum(axis=1)
                    df[mean_col_name] = df[included_vars].mean(axis=1)
                    total_created_vars += 2

                    detailed_results.append({
                        'group': group_name,
                        'variables': len(included_vars),
                        'reverse_count': len(reverse_vars),
                        'sum_col': sum_col_name,
                        'mean_col': mean_col_name,
                        'range': f"{group_min}~{group_max}"
                    })

                except Exception as e:
                    print(f"계산 오류 {group_name}: {e}")

            completed_groups += 1
            import time
            time.sleep(0.1)

        # 완료 메시지
        result_message = f"""
🎉 고급 계산 완료!

📊 처리 결과:
  • 처리된 그룹: {completed_groups}개
  • 생성된 변수: {total_created_vars}개
  • 사용자 지정 역코딩: {sum(r['reverse_count'] for r in detailed_results)}개

📋 생성된 변수들:
"""

        for result in detailed_results:
            result_message += f"  🎯 {result['group']} ({result['range']}): {result['sum_col']}, {result['mean_col']}\n"

        result_message += f"\n💾 '결과 저장' 버튼으로 엑셀에 저장하세요."

        # 메인 리스트 업데이트
        refresh_main_variable_list()

        messagebox.showinfo("고급 계산 완료!", f"✅ {completed_groups}개 그룹 완료!\n✨ 생성: {total_created_vars}개 변수\n🔄 역코딩: {sum(r['reverse_count'] for r in detailed_results)}개")
        update_result_text(result_message)

    except Exception as e:
        messagebox.showerror("오류", f"고급 계산 중 오류: {str(e)}")
        update_result_text(f"❌ 고급 계산 실패\n오류: {str(e)}")


def reset_analysis():
    """분석 초기화"""
    global df, original_column_order, reverse_coded_columns, selected_factor_vars

    if df is None:
        return

    # 초기화 확인 메시지
    result = messagebox.askyesno("분석 초기화 확인",
                                "⚠️ 정말로 모든 분석을 초기화하시겠습니까?\n\n"
                                "다음 내용이 모두 삭제됩니다:\n"
                                "• 계산된 모든 합계/평균 변수\n"
                                "• 생성된 모든 역코딩 변수\n"
                                "• 모든 변수 선택 상태\n\n"
                                "원본 엑셀 파일 상태로 완전히 돌아갑니다.")

    if not result:
        return

    try:
        file_path = entry_file_path.get()
        if file_path:
            df = pd.read_excel(file_path)
            original_column_order = list(df.columns)
            reverse_coded_columns = {}
            selected_factor_vars = []

            listbox_factor_vars.delete(0, tk.END)
            for col in df.columns:
                listbox_factor_vars.insert(tk.END, col)

            # 워크플로우 상태 완전 초기화
            workflow_state['step'] = 1
            workflow_state['file_loaded'] = True
            workflow_state['variables_selected'] = False
            workflow_state['reverse_coding_done'] = False
            workflow_state['calculation_done'] = False
            update_button_states()

            update_result_text("🔄 모든 분석이 초기화되었습니다!\n\n📋 1단계: 첫 번째 요인에 속하는 변수들을 선택하세요\n\n💡 원본 엑셀 파일 상태로 완전히 돌아갔습니다.")
            messagebox.showinfo("초기화 완료", "✅ 모든 분석이 초기화되어 원본 상태로 돌아갔습니다!")
    except Exception as e:
        messagebox.showerror("오류", f"초기화 중 오류가 발생했습니다:\n{e}")


def apply_text_formatting():
    """텍스트에 색상 및 강조 효과 적용"""
    # 태그 설정
    text_result.tag_configure("header", font=("Arial", 16, "bold"), foreground="#2E86AB")
    text_result.tag_configure("success", font=("Arial", 15, "bold"), foreground="#008000")
    text_result.tag_configure("warning", font=("Arial", 14, "bold"), foreground="#FF6600")
    text_result.tag_configure("info", font=("Arial", 13, "bold"), foreground="#4682B4")
    text_result.tag_configure("variable", font=("Arial", 14, "bold"), foreground="#8B4513", underline=True)
    text_result.tag_configure("number", font=("Arial", 14, "bold"), foreground="#DC143C")

    content = text_result.get(1.0, tk.END)

    # 패턴별 강조 적용
    import re

    # 헤더 박스 강조
    for match in re.finditer(r'┏.*?┓', content, re.DOTALL):
        start_idx = f"1.0 + {match.start()}c"
        end_idx = f"1.0 + {match.end()}c"
        text_result.tag_add("header", start_idx, end_idx)

    # 완료 메시지 강조
    for match in re.finditer(r'✅[^\\n]*', content):
        start_idx = f"1.0 + {match.start()}c"
        end_idx = f"1.0 + {match.end()}c"
        text_result.tag_add("success", start_idx, end_idx)

    # 경고 메시지 강조
    for match in re.finditer(r'⚠️[^\\n]*|🚨[^\\n]*', content):
        start_idx = f"1.0 + {match.start()}c"
        end_idx = f"1.0 + {match.end()}c"
        text_result.tag_add("warning", start_idx, end_idx)

    # 변수명 강조 (▶️ [숫자] 변수명 패턴)
    for match in re.finditer(r'▶️ \[\d+\] ([^\s]+)', content):
        var_start = match.start(1)
        var_end = match.end(1)
        start_idx = f"1.0 + {var_start}c"
        end_idx = f"1.0 + {var_end}c"
        text_result.tag_add("variable", start_idx, end_idx)

    # 숫자 강조 (개수, 진행률 등)
    for match in re.finditer(r'\((\d+)개\)|\((\d+)/(\d+)\)', content):
        start_idx = f"1.0 + {match.start()}c"
        end_idx = f"1.0 + {match.end()}c"
        text_result.tag_add("number", start_idx, end_idx)


def update_result_text(text):
    """결과 텍스트 업데이트 (읽기전용 모드 처리)"""
    if text_result is None:
        print(f"텍스트 업데이트 실패: {text}")
        return

    try:
        text_result.config(state=tk.NORMAL)  # 임시로 수정 가능하게
        text_result.delete(1.0, tk.END)
        text_result.insert(tk.END, text)

        # 텍스트 포맷팅 적용
        apply_text_formatting()

        text_result.config(state=tk.DISABLED)  # 다시 읽기전용으로
    except Exception as e:
        print(f"텍스트 업데이트 오류: {e}")


def show_current_selection():
    """현재 선택된 변수들을 실시간으로 표시하고 계산 준비 상태도 업데이트"""
    selected_indices = listbox_factor_vars.curselection()
    selected_vars = [listbox_factor_vars.get(idx) for idx in selected_indices]

    # 전체 텍스트를 가져와서 선택 정보 부분만 정확히 교체
    text_result.config(state=tk.NORMAL)
    full_text = text_result.get(1.0, tk.END)

    # 기본 안내 메시지가 있는지 확인
    if "환영합니다!" in full_text or not full_text.strip():
        # 기본 안내 메시지 표시
        base_content = "환영합니다!\n\n1. 먼저 엑셀 파일을 선택하세요\n2. 같은 요인 변수들을 선택하세요\n3. 필요시 역코딩 실행\n4. 합계/평균 계산\n\n선택된 변수들이 아래에 표시됩니다."
    else:
        # 기존 내용 유지하되 선택 정보와 계산 준비 정보는 업데이트
        lines = full_text.split('\n')
        result_lines = []
        skip_next_lines = False

        for line in lines:
            # 선택 정보 블록 시작
            if "🎯 선택된 변수" in line:
                skip_next_lines = True
                continue
            # 계산 준비 정보 블록 시작
            elif "✅ 계산 준비 완료" in line or "합계/평균 계산에 사용될 변수들:" in line:
                skip_next_lines = True
                continue
            # 스킵 중인 블록의 내용들
            elif skip_next_lines and (line.startswith("  [") or line.strip() == "" or "현황:" in line or "다음 단계:" in line):
                continue
            else:
                skip_next_lines = False
                result_lines.append(line)

        base_content = '\n'.join(result_lines).rstrip()

    # 선택된 변수들이 있으면 계산 준비 상태만 표시 (중복 제거)
    if selected_vars:
        # 계산 준비 상태만 표시
        calculation_ready_block = f"""

✅ 계산 준비 완료

합계/평균 계산에 사용될 변수들:
"""
        for i, var in enumerate(selected_vars):
            if var.startswith('역_'):
                calculation_ready_block += f"  [{i+1}] {var} (역코딩) (더블클릭으로 제거)\n"
            else:
                calculation_ready_block += f"  [{i+1}] {var} (원본) (더블클릭으로 제거)\n"

        reverse_count = len([v for v in selected_vars if v.startswith('역_')])
        original_count = len(selected_vars) - reverse_count

        calculation_ready_block += f"""
현황: 원본 {original_count}개, 역코딩 {reverse_count}개, 총 {len(selected_vars)}개

다음 단계: "합계 및 평균 계산" 버튼 클릭
"""

        final_text = base_content + calculation_ready_block
    else:
        final_text = base_content

    # 전체 텍스트 교체
    text_result.delete(1.0, tk.END)
    text_result.insert(tk.END, final_text)

    # 포맷팅 적용
    apply_text_formatting()

    text_result.config(state=tk.DISABLED)


def show_final_variable_summary():
    """3단계 전 최종 선택된 변수들 상태 요약 표시"""
    global selected_factor_vars, reverse_coded_columns

    if not selected_factor_vars:
        messagebox.showerror("오류", "먼저 요인 변수들을 선택하세요!")
        return

    summary_text = f"""
╔══════════════════════════════════════════════════════════════════════╗
║  📋 ✨ 3단계: 합계/평균 계산 준비 상태 ✨                           ║
╚══════════════════════════════════════════════════════════════════════╝

🎯 최종 선택된 요인 변수들:
"""

    for i, var in enumerate(selected_factor_vars):
        if var in reverse_coded_columns:
            # 역코딩된 변수
            reverse_var = reverse_coded_columns[var]
            summary_text += f"  ▶️ [{i+1}] {var} → {reverse_var} (🔄 역코딩 데이터 사용됨)\n"
        else:
            # 원본 변수
            summary_text += f"  ▶️ [{i+1}] {var} (📊 원본 데이터 사용됨)\n"

    reverse_count = len([v for v in selected_factor_vars if v in reverse_coded_columns])
    original_count = len(selected_factor_vars) - reverse_count

    summary_text += f"""
════════════════════════════════════════════════════════

📈 변수 현황:
  📊 원본 데이터: {original_count}개
  🔄 역코딩 데이터: {reverse_count}개
  📋 총 변수: {len(selected_factor_vars)}개

🧮 계산될 항목:
  🔢 [요인명]_합계 = {len(selected_factor_vars)}개 변수의 합
  📈 [요인명]_평균 = {len(selected_factor_vars)}개 변수의 평균

⚡ 준비 완료! 이제 "합계 및 평균 계산" 버튼을 클릭하세요! ⚡
"""

    update_result_text(summary_text)
    root.update_idletasks()


# show_calculation_ready_summary 함수는 더 이상 사용하지 않음
# show_current_selection에서 실시간으로 계산 준비 상태를 표시함


def refresh_main_variable_list():
    """메인 화면의 변수 리스트를 현재 데이터프레임 기준으로 새로고침"""
    global df

    # 데이터 및 GUI 컴포넌트 검증
    if df is None:
        print("데이터프레임이 None입니다.")
        return

    if listbox_factor_vars is None:
        print("리스트박스가 초기화되지 않았습니다.")
        return

    try:
        # 현재 선택된 인덱스들 저장
        current_selected_vars = []
        try:
            for idx in listbox_factor_vars.curselection():
                var_name = listbox_factor_vars.get(idx)
                if var_name:  # 빈 문자열 체크
                    current_selected_vars.append(var_name)
        except tk.TclError as e:
            print(f"선택 상태 저장 중 오류: {e}")

        # 리스트박스 내용 업데이트
        listbox_factor_vars.delete(0, tk.END)

        # 데이터프레임 컬럼 검증 후 추가
        if hasattr(df, 'columns') and len(df.columns) > 0:
            for col in df.columns:
                if col is not None and str(col).strip():  # 유효한 컬럼명 체크
                    listbox_factor_vars.insert(tk.END, str(col))
        else:
            print("유효한 컬럼이 없습니다.")
            return

        # 이전 선택 복원 (변수명 기준)
        for i, var in enumerate(df.columns):
            if str(var) in current_selected_vars:
                try:
                    listbox_factor_vars.selection_set(i)
                except tk.TclError as e:
                    print(f"선택 복원 중 오류: {e}")

        # 선택 상태 업데이트 (안전하게)
        if root is not None:
            root.after(50, lambda: show_current_selection() if show_current_selection else None)

    except Exception as e:
        print(f"변수 리스트 새로고침 중 오류: {e}")


def refresh_main_variable_list_with_selection(selected_vars):
    """메인 화면의 변수 리스트를 새로고침하면서 특정 변수들 선택 유지"""
    global df

    if df is not None:
        # 리스트박스 내용 업데이트
        listbox_factor_vars.delete(0, tk.END)
        for col in df.columns:
            listbox_factor_vars.insert(tk.END, col)

        # 지정된 변수들 선택 복원 (변수명 기준)
        for i, var in enumerate(df.columns):
            if var in selected_vars:
                listbox_factor_vars.selection_set(i)

        # 선택 상태 업데이트
        root.after(50, show_current_selection)


# GUI 안전 초기화
try:
    root = tk.Tk()
    root.title("🔢 변수계산 및 역코딩 최종 프로그램")
    root.geometry("1200x800")
    root.configure(bg=COLORS['light'])

    # 프로그램 종료 시 안전하게 처리
    def on_closing():
        try:
            root.destroy()
        except:
            pass

    root.protocol("WM_DELETE_WINDOW", on_closing)

except Exception as e:
    print(f"GUI 초기화 오류: {e}")
    exit(1)

# 스타일 설정
style = ttk.Style()
style.theme_use('clam')

# 메인 제목
title_frame = tk.Frame(root, bg=COLORS['primary'], height=80)
title_frame.pack(fill=tk.X, padx=10, pady=10)
title_frame.pack_propagate(False)

tk.Label(title_frame, text="🔢 변수계산 및 역코딩 최종 프로그램",
         font=("Arial", 18, "bold"), fg=COLORS['dark'],
         bg=COLORS['primary']).pack(expand=True)

# 파일 선택 영역
file_frame = tk.Frame(root, bg=COLORS['white'], relief=tk.RAISED, bd=2)
file_frame.pack(fill=tk.X, padx=10, pady=5)

file_inner = tk.Frame(file_frame, bg=COLORS['white'])
file_inner.pack(fill=tk.X, padx=15, pady=10)

tk.Label(file_inner, text="📁 엑셀 파일:", font=("Arial", 11, "bold"),
         fg=COLORS['dark'], bg=COLORS['white']).pack(side=tk.LEFT)

# 전역변수에 안전하게 할당
try:
    entry_file_path = tk.Entry(file_inner, width=60, font=("Arial", 10))
    entry_file_path.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
    btn_browse = tk.Button(file_inner, text="파일 선택", command=select_file,
                          bg=COLORS['primary'], fg=COLORS['button_text'], font=("Arial", 10, "bold"),
                          padx=20, pady=5)
    btn_browse.pack(side=tk.RIGHT)
except Exception as e:
    print(f"파일 선택 GUI 생성 오류: {e}")
    messagebox.showerror("오류", "GUI 생성 중 오류가 발생했습니다.")

# 메인 콘텐츠 영역 (좌우 분할)
main_frame = tk.Frame(root, bg=COLORS['light'])
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

# 왼쪽 영역 (변수 선택)
left_frame = tk.Frame(main_frame, bg=COLORS['white'], relief=tk.RAISED, bd=2)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

# 왼쪽 제목
left_title = tk.Frame(left_frame, bg=COLORS['secondary'], height=50)
left_title.pack(fill=tk.X, padx=5, pady=5)
left_title.pack_propagate(False)
tk.Label(left_title, text="📋 1단계: 요인 변수 선택",
         font=("Arial", 12, "bold"), fg=COLORS['dark'],
         bg=COLORS['secondary']).pack(expand=True)

# 변수 선택 리스트
list_frame = tk.Frame(left_frame, bg=COLORS['white'])
list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

tk.Label(list_frame, text="같은 요인에 속하는 변수들을 선택하세요\n• 클릭: 개별 선택/해제 • Shift+클릭: 비슷한 변수 그룹 토글 (모두선택↔모두해제)",
         font=("Arial", 12), fg=COLORS['dark'], bg=COLORS['white'],
         justify=tk.LEFT).pack(anchor=tk.W)

# 리스트박스 안전 생성
try:
    listbox_factor_vars = tk.Listbox(list_frame, selectmode=tk.MULTIPLE,
                                    font=("Arial", 14), height=15,
                                    bg=COLORS['white'], fg=COLORS['dark'],
                                    selectbackground=COLORS['info'],
                                    activestyle='dotbox')
    listbox_factor_vars.pack(fill=tk.BOTH, expand=True, pady=5)
except Exception as e:
    print(f"리스트박스 생성 오류: {e}")
    messagebox.showerror("오류", "리스트박스 생성 중 오류가 발생했습니다.")

# 클릭 선택 기능 (드래그 제거) - 성능 최적화
def on_click(event):
    try:
        widget = event.widget
        index = widget.nearest(event.y)

        if index < 0 or index >= widget.size():
            return "break"

        # 클릭: 토글 방식 (기존 선택 유지)
        current_selection = list(widget.curselection())
        if index in current_selection:
            widget.selection_clear(index)
        else:
            widget.selection_set(index)

        # 선택 상태 실시간 업데이트 (디바운싱으로 성능 최적화)
        if hasattr(on_click, '_update_timer'):
            root.after_cancel(on_click._update_timer)
        on_click._update_timer = root.after(50, show_current_selection)

        return "break"
    except Exception as e:
        print(f"클릭 이벤트 처리 중 오류: {e}")
        return "break"

# Shift+클릭으로 비슷한 변수들 자동 선택/해제 (토글) - 성능 최적화
def on_shift_click(event):
    try:
        widget = event.widget
        index = widget.nearest(event.y)

        if index < 0 or index >= widget.size():
            return "break"

        clicked_var = widget.get(index)
        if not clicked_var:
            return "break"

        # 현재 리스트의 모든 변수들 가져오기 (캐싱으로 성능 개선)
        all_vars = [widget.get(i) for i in range(widget.size())]

        # 비슷한 변수들 찾기
        similar_vars = find_similar_variables(clicked_var, all_vars)

        if not similar_vars:
            return "break"

        # 현재 선택 상태 확인
        current_selection = list(widget.curselection())
        similar_indices = []
        for i, var in enumerate(all_vars):
            if var in similar_vars:
                similar_indices.append(i)

        if not similar_indices:
            return "break"

        # 비슷한 변수들이 모두 선택되어 있는지 확인
        all_selected = all(i in current_selection for i in similar_indices)

        # 배치 처리로 성능 향상
        if all_selected:
            # 모두 선택되어 있으면 → 모두 해제
            for i in similar_indices:
                widget.selection_clear(i)
        else:
            # 일부만 선택되어 있거나 선택 안되어 있으면 → 모두 선택
            for i in similar_indices:
                widget.selection_set(i)

        # 선택 상태 실시간 업데이트 (디바운싱)
        if hasattr(on_shift_click, '_update_timer'):
            root.after_cancel(on_shift_click._update_timer)
        on_shift_click._update_timer = root.after(50, show_current_selection)

        return "break"
    except Exception as e:
        print(f"Shift+클릭 이벤트 처리 중 오류: {e}")
        return "break"

# 모든 기본 선택 이벤트 비활성화 후 커스텀 이벤트만 활성화
def disable_default_selection(_):
    return "break"

# 기본 이벤트들 모두 차단
listbox_factor_vars.bind("<Button-1>", disable_default_selection)
listbox_factor_vars.bind("<ButtonRelease-1>", disable_default_selection)
listbox_factor_vars.bind("<B1-Motion>", disable_default_selection)
listbox_factor_vars.bind("<Double-Button-1>", disable_default_selection)

# 커스텀 이벤트만 허용
listbox_factor_vars.bind("<Button-1>", on_click)
listbox_factor_vars.bind("<Shift-Button-1>", on_shift_click)

# 왼쪽 버튼들
left_button_frame = tk.Frame(left_frame, bg=COLORS['white'])
left_button_frame.pack(fill=tk.X, padx=10, pady=10)

# 왼쪽 버튼들을 전역변수에 할당
try:
    btn_select_factor = tk.Button(left_button_frame, text="✅ 요인 변수 선택 완료",
                                 command=select_factor_variables,
                                 bg=COLORS['primary'], fg=COLORS['button_text'],
                                 font=("Arial", 11, "bold"), pady=8)
    btn_select_factor.pack(fill=tk.X, pady=2)

    btn_reverse = tk.Button(left_button_frame, text="🔄 역코딩할 변수 선택",
                           command=show_reverse_coding_dialog,
                           bg=COLORS['success'], fg=COLORS['button_text'],
                           font=("Arial", 11, "bold"), pady=8)
    btn_reverse.pack(fill=tk.X, pady=2)

    btn_calculate = tk.Button(left_button_frame, text="📊 합계 및 평균 계산",
                             command=calculate_factor_statistics,
                             bg=COLORS['info'], fg=COLORS['button_text'],
                             font=("Arial", 11, "bold"), pady=8)
    btn_calculate.pack(fill=tk.X, pady=2)
except Exception as e:
    print(f"왼쪽 버튼 생성 오류: {e}")
    messagebox.showerror("오류", "버튼 생성 중 오류가 발생했습니다.")

# 오른쪽 영역 (처리 결과 전체)
right_frame = tk.Frame(main_frame, bg=COLORS['white'], relief=tk.RAISED, bd=2)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

# 처리 결과 제목
result_title = tk.Frame(right_frame, bg=COLORS['secondary'], height=60)
result_title.pack(fill=tk.X, padx=5, pady=5)
result_title.pack_propagate(False)
tk.Label(result_title, text="📊 처리 결과 및 진행 상황",
         font=("Arial", 14, "bold"), fg=COLORS['dark'],
         bg=COLORS['secondary']).pack(expand=True)

# 사용법 안내
usage_frame = tk.Frame(right_frame, bg=COLORS['info'], height=40)
usage_frame.pack(fill=tk.X, padx=5, pady=(0, 5))
usage_frame.pack_propagate(False)
tk.Label(usage_frame, text="💡 변수명 더블클릭으로 선택 해제 가능",
         font=("Arial", 11, "bold"), fg=COLORS['dark'],
         bg=COLORS['info']).pack(expand=True)

# 처리 결과 텍스트 안전 생성
try:
    text_result = tk.Text(right_frame, width=50, font=("Arial", 14, "bold"),
                         bg=COLORS['light'], fg=COLORS['dark'], wrap=tk.WORD,
                         state=tk.DISABLED, padx=10, pady=10)
    text_result.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
except Exception as e:
    print(f"텍스트 결과창 생성 오류: {e}")
    messagebox.showerror("오류", "결과창 생성 중 오류가 발생했습니다.")

# 처리 결과창에서 변수 삭제 기능 (개선된 변수명 추출)
def on_result_double_click(event):
    """처리 결과창에서 변수 더블클릭 시 선택에서 제거"""
    global is_processing_click

    try:
        # 이미 처리 중이면 무시
        if is_processing_click:
            return "break"

        # 처리 시작 플래그 설정
        is_processing_click = True

        # 현재 커서 위치의 줄 가져오기
        text_result.config(state=tk.NORMAL)

        # 클릭한 위치의 줄 찾기
        click_index = text_result.index("@%s,%s" % (event.x, event.y))
        line_start = text_result.index("%s linestart" % click_index)
        line_end = text_result.index("%s lineend" % click_index)
        current_line = text_result.get(line_start, line_end)

        var_name = None

        # 변수명 추출 (새로운 형식에서) - 개선된 로직
        if "[" in current_line and "] " in current_line and "(더블클릭으로 제거)" in current_line:
            try:
                # "[1] 변수명 (역코딩) (더블클릭으로 제거)" 또는 "[1] 변수명 (원본) (더블클릭으로 제거)" 형식에서 추출
                start_pos = current_line.find("] ") + 2

                if start_pos > 1:  # "] " 문자열이 실제로 발견된 경우
                    # (역코딩) 또는 (원본) 앞까지 추출
                    temp_text = current_line[start_pos:]
                    extracted_text = ""

                    if " (역코딩)" in temp_text:
                        end_pos = temp_text.find(" (역코딩)")
                        extracted_text = temp_text[:end_pos].strip()
                    elif " (원본)" in temp_text:
                        end_pos = temp_text.find(" (원본)")
                        extracted_text = temp_text[:end_pos].strip()
                    else:
                        # 기존 방식 (이전 형식 호환)
                        end_pos = current_line.find(" (더블클릭으로 제거)")
                        if end_pos > start_pos:
                            extracted_text = current_line[start_pos:end_pos].strip()

                    if extracted_text:
                        # 실제 변수 리스트에서 정확한 매칭 찾기
                        for i in range(listbox_factor_vars.size()):
                            list_var = listbox_factor_vars.get(i)
                            if list_var == extracted_text:
                                var_name = list_var
                                break

                        # 정확한 매칭이 없으면 디버깅 정보 출력
                        if not var_name:
                            print(f"DEBUG: 추출된 변수명 '{extracted_text}'를 리스트에서 찾을 수 없음")
                            print(f"DEBUG: 현재 리스트의 변수들: {[listbox_factor_vars.get(i) for i in range(listbox_factor_vars.size())]}")

            except Exception as e:
                print(f"변수명 추출 중 오류: {e}")

        elif "📊 " in current_line:
            try:
                # "📊 변수명 → 역_변수명" 형식에서 추출
                parts = current_line.split("📊 ")[1]
                if " " in parts:
                    var_name = parts.split(" ")[0].strip()
                else:
                    var_name = parts.strip()
            except Exception as e:
                print(f"📊 형식 처리 중 오류: {e}")

        elif "• " in current_line:
            try:
                # "• 변수명 (설명)" 형식에서 추출
                var_name = current_line.split("• ")[1].split(" ")[0].strip()
            except Exception as e:
                print(f"• 형식 처리 중 오류: {e}")

        if var_name:
            # 시각적 하이라이트
            try:
                text_result.tag_configure("delete_highlight", background="#FF6B6B", foreground="white")
                text_result.tag_add("delete_highlight", line_start, line_end)
                text_result.config(state=tk.DISABLED)
            except Exception as e:
                print(f"하이라이트 설정 중 오류: {e}")

            def complete_deletion():
                global is_processing_click

                try:
                    # 하이라이트 제거
                    text_result.config(state=tk.NORMAL)
                    text_result.tag_remove("delete_highlight", line_start, line_end)
                    text_result.config(state=tk.DISABLED)

                    # 메인 리스트에서 해당 변수 선택 해제
                    removed = False
                    current_selection = list(listbox_factor_vars.curselection())

                    for i in range(listbox_factor_vars.size()):
                        if listbox_factor_vars.get(i) == var_name:
                            # 해당 변수가 실제로 선택되어 있는지 확인
                            if i in current_selection:
                                listbox_factor_vars.selection_clear(i)
                                removed = True
                            else:
                                # 선택되어 있지 않다면 조용히 처리 (메시지 없음)
                                removed = True  # 처리 완료로 간주
                            break

                    if not removed:
                        messagebox.showwarning("경고", f"변수 '{var_name}'을 찾을 수 없습니다.")
                    else:
                        # 선택 상태 업데이트 (지연을 늘려서 확실히 반영되도록)
                        root.after(50, show_current_selection)

                except Exception as e:
                    print(f"삭제 처리 중 오류: {e}")
                finally:
                    # 처리 완료 플래그 해제
                    is_processing_click = False

            # 0.2초 후에 삭제 완료
            root.after(200, complete_deletion)

        else:
            text_result.config(state=tk.DISABLED)
            # 처리할 변수가 없으면 즉시 플래그 해제
            is_processing_click = False

    except Exception as e:
        print(f"더블클릭 처리 중 전체 오류: {e}")
        # 오류 발생 시 플래그 해제
        is_processing_click = False
        try:
            text_result.config(state=tk.DISABLED)
        except:
            pass

    return "break"

text_result.bind("<Double-Button-1>", on_result_double_click)


# 하단 버튼 영역
bottom_frame = tk.Frame(root, bg=COLORS['light'])
bottom_frame.pack(fill=tk.X, padx=10, pady=10)

# 하단 버튼들을 전역변수에 할당
try:
    btn_save = tk.Button(bottom_frame, text="💾 결과 저장", command=save_to_excel,
                        bg=COLORS['success'], fg=COLORS['button_text'], font=("Arial", 11, "bold"),
                        padx=30, pady=10)
    btn_save.pack(side=tk.LEFT)

    btn_prepare_next = tk.Button(bottom_frame, text="🚀 다음 요인 계산 준비", command=prepare_next_factor,
                               bg=COLORS['info'], fg=COLORS['button_text'], font=("Arial", 11, "bold"),
                               padx=30, pady=10)
    btn_prepare_next.pack(side=tk.LEFT, padx=10)

    # 빠른 계산 버튼 추가 (중앙에 강조, 더 눈에 띄게)
    btn_quick_calc = tk.Button(bottom_frame, text="⚡ 빠른 계산", command=quick_calculation,
                              bg=COLORS['highlight'], fg=COLORS['button_text'], font=("Arial", 13, "bold"),
                              padx=50, pady=15, relief=tk.RAISED, bd=4,
                              activebackground=COLORS['glow'], activeforeground=COLORS['button_text'])
    btn_quick_calc.pack(side=tk.LEFT, padx=25)

    # 빠른 계산 버튼에 호버 효과 추가
    def on_quick_enter(event):
        btn_quick_calc.config(bg=COLORS['glow'])

    def on_quick_leave(event):
        btn_quick_calc.config(bg=COLORS['highlight'])

    btn_quick_calc.bind("<Enter>", on_quick_enter)
    btn_quick_calc.bind("<Leave>", on_quick_leave)

    btn_reset = tk.Button(bottom_frame, text="🔄 분석 초기화\n(모든 계산 초기화)", command=reset_analysis,
                         bg=COLORS['warning'], fg=COLORS['button_text'], font=("Arial", 10, "bold"),
                         padx=20, pady=10)
    btn_reset.pack(side=tk.RIGHT)

    # 분석 초기화 버튼 툴팁 추가
    def show_reset_tooltip(event):
        import tkinter.messagebox as msg
        msg.showinfo("분석 초기화 안내",
                    "⚠️ 분석 초기화 기능 안내\n\n"
                    "• 지금까지 계산된 모든 합계/평균 변수가 삭제됩니다\n"
                    "• 모든 역코딩 변수가 삭제됩니다\n"
                    "• 원본 엑셀 파일 상태로 완전히 돌아갑니다\n"
                    "• 모든 선택 상태가 초기화됩니다\n\n"
                    "💡 일부 변수만 다시 계산하려면 '다음 요인 계산 준비'를 사용하세요")

    # 우클릭으로 도움말 표시
    btn_reset.bind("<Button-2>", show_reset_tooltip)  # 맥: Command+클릭
    btn_reset.bind("<Button-3>", show_reset_tooltip)  # 윈도우: 우클릭
except Exception as e:
    print(f"하단 버튼 생성 오류: {e}")
    messagebox.showerror("오류", "버튼 생성 중 오류가 발생했습니다.")

# 초기화 완료 검증
def verify_initialization():
    """GUI 컴포넌트 초기화 검증"""
    components = {
        'root': root,
        'entry_file_path': entry_file_path,
        'listbox_factor_vars': listbox_factor_vars,
        'text_result': text_result,
        'btn_browse': btn_browse,
        'btn_select_factor': btn_select_factor,
        'btn_reverse': btn_reverse,
        'btn_calculate': btn_calculate,
        'btn_prepare_next': btn_prepare_next
    }

    missing_components = []
    for name, component in components.items():
        if component is None:
            missing_components.append(name)

    if missing_components:
        error_msg = f"다음 GUI 컴포넌트가 초기화되지 않았습니다: {', '.join(missing_components)}"
        print(error_msg)
        if root:
            messagebox.showerror("초기화 오류", error_msg)
        return False

    return True

# 초기화 검증 후 프로그램 시작
try:
    if verify_initialization():
        # 초기 표시 업데이트
        update_result_text("환영합니다!\n\n1. 먼저 엑셀 파일을 선택하세요\n2. 같은 요인 변수들을 선택하세요\n3. 필요시 역코딩 실행\n4. 합계/평균 계산\n\n선택된 변수들이 아래에 표시됩니다.")

        # 초기 버튼 상태 설정 (약간의 지연 후)
        workflow_state['step'] = 1
        print(f"초기 단계 설정: {workflow_state}")

        # 즉시 한 번 시도하고, 실패하면 지연 후 재시도
        try:
            update_button_states()
        except:
            print("초기 버튼 상태 설정 실패, 지연 후 재시도")
            root.after(500, update_button_states)  # GUI 완전히 생성된 후 호출

        print("프로그램이 성공적으로 초기화되었습니다.")

        # GUI 실행
        root.mainloop()
    else:
        print("프로그램 초기화에 실패했습니다.")

except Exception as e:
    print(f"프로그램 실행 중 오류: {e}")
    if root:
        try:
            root.destroy()
        except:
            pass
"""
설문지 매핑 프로그램 - 고급 사용자 친화 버전
============================================
주요 기능:
- 드래그 앤 드롭 파일 업로드
- 실시간 데이터 미리보기
- 배치 매핑 (동일한 선택지를 여러 변수에 일괄 적용)
- 자동 저장 및 되돌리기
- 진행률 표시
- 매핑 템플릿 저장/불러오기
- 검색 및 필터링
- 단축키 지원
"""

import pandas as pd
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import traceback
from datetime import datetime

class SurveyMappingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("설문지 매핑 프로그램 - 고급 버전")
        self.root.geometry("1400x900")

        # 데이터 초기화
        self.df = None
        self.original_df = None  # 백업용
        self.file_path = None
        self.user_defined_mappings = {}
        self.mapping_history = []  # 되돌리기용
        self.skip_values = set()
        self.current_mapping_session = {}

        # 색상 테마
        self.colors = {
            'primary': '#2196F3',
            'secondary': '#4CAF50',
            'warning': '#FF9800',
            'danger': '#F44336',
            'success': '#4CAF50',
            'bg_dark': '#37474F',
            'bg_light': '#ECEFF1',
            'text_dark': '#212121',
            'text_light': '#FFFFFF'
        }

        self.setup_ui()
        self.setup_shortcuts()

    def setup_shortcuts(self):
        """단축키 설정"""
        self.root.bind('<Control-o>', lambda e: self.select_file())
        self.root.bind('<Control-s>', lambda e: self.save_to_excel())
        self.root.bind('<Control-z>', lambda e: self.undo_mapping())
        self.root.bind('<F5>', lambda e: self.refresh_preview())
        self.root.bind('<Control-f>', lambda e: self.focus_search())

    def setup_ui(self):
        """UI 구성"""
        # 메인 컨테이너
        main_container = tk.Frame(self.root, bg=self.colors['bg_light'])
        main_container.pack(fill=tk.BOTH, expand=True)

        # 상단 툴바
        self.create_toolbar(main_container)

        # 메인 영역 (좌우 분할)
        paned = tk.PanedWindow(main_container, orient=tk.HORIZONTAL, bg=self.colors['bg_light'])
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 좌측 패널 (변수 목록 및 제어)
        left_panel = tk.Frame(paned, bg='white', relief=tk.RAISED, borderwidth=1)
        paned.add(left_panel, width=450)
        self.create_left_panel(left_panel)

        # 우측 패널 (데이터 미리보기 및 매핑 정보)
        right_panel = tk.Frame(paned, bg='white', relief=tk.RAISED, borderwidth=1)
        paned.add(right_panel)
        self.create_right_panel(right_panel)

        # 하단 상태바
        self.create_statusbar(main_container)

    def create_toolbar(self, parent):
        """상단 툴바 생성"""
        toolbar = tk.Frame(parent, bg=self.colors['bg_dark'], height=60)
        toolbar.pack(fill=tk.X)
        toolbar.pack_propagate(False)

        # 파일 업로드 영역
        upload_frame = tk.Frame(toolbar, bg=self.colors['bg_dark'])
        upload_frame.pack(side=tk.LEFT, padx=20, pady=10)

        tk.Label(upload_frame, text="📁 파일:", bg=self.colors['bg_dark'],
                fg=self.colors['text_light'], font=('맑은 고딕', 10, 'bold')).pack(side=tk.LEFT, padx=5)

        self.entry_file_path = tk.Entry(upload_frame, width=40, font=('맑은 고딕', 9))
        self.entry_file_path.pack(side=tk.LEFT, padx=5)
        self.entry_file_path.drop_target_register(DND_FILES)
        self.entry_file_path.dnd_bind('<<Drop>>', self.drop_file)

        btn_browse = tk.Button(upload_frame, text="찾아보기", command=self.select_file,
                              bg=self.colors['primary'], fg='white', font=('맑은 고딕', 9, 'bold'),
                              relief=tk.FLAT, padx=15, pady=5, cursor='hand2')
        btn_browse.pack(side=tk.LEFT, padx=5)

        # 주요 작업 버튼들
        button_frame = tk.Frame(toolbar, bg=self.colors['bg_dark'])
        button_frame.pack(side=tk.RIGHT, padx=20, pady=10)

        buttons = [
            ("🔄 되돌리기", self.undo_mapping, self.colors['warning']),
            ("💾 저장", self.save_to_excel, self.colors['success']),
            ("📄 템플릿 저장", self.save_template, self.colors['secondary']),
            ("📂 템플릿 불러오기", self.load_template, self.colors['secondary'])
        ]

        for text, command, color in buttons:
            btn = tk.Button(button_frame, text=text, command=command,
                          bg=color, fg='white', font=('맑은 고딕', 9, 'bold'),
                          relief=tk.FLAT, padx=10, pady=5, cursor='hand2')
            btn.pack(side=tk.LEFT, padx=3)
            self.add_hover_effect(btn, color)

    def create_left_panel(self, parent):
        """좌측 패널 생성 (변수 목록 및 제어)"""
        # 제목
        title_frame = tk.Frame(parent, bg=self.colors['primary'], height=40)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        tk.Label(title_frame, text="📋 변수 목록 및 매핑 제어",
                bg=self.colors['primary'], fg='white',
                font=('맑은 고딕', 12, 'bold')).pack(pady=8)

        # 검색 바
        search_frame = tk.Frame(parent, bg='white')
        search_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(search_frame, text="🔍", font=('맑은 고딕', 12), bg='white').pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_columns)
        search_entry = tk.Entry(search_frame, textvariable=self.search_var,
                               font=('맑은 고딕', 10), relief=tk.FLAT,
                               bg=self.colors['bg_light'])
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 통계 정보
        self.stats_frame = tk.LabelFrame(parent, text="📊 파일 정보",
                                         font=('맑은 고딕', 10, 'bold'),
                                         bg='white', fg=self.colors['text_dark'])
        self.stats_frame.pack(fill=tk.X, padx=10, pady=5)

        self.lbl_total_vars = tk.Label(self.stats_frame, text="총 변수: 0",
                                       bg='white', font=('맑은 고딕', 9))
        self.lbl_total_vars.pack(anchor='w', padx=10, pady=2)

        self.lbl_mapped_vars = tk.Label(self.stats_frame, text="매핑 완료: 0",
                                        bg='white', font=('맑은 고딕', 9))
        self.lbl_mapped_vars.pack(anchor='w', padx=10, pady=2)

        self.lbl_total_rows = tk.Label(self.stats_frame, text="총 행: 0",
                                       bg='white', font=('맑은 고딕', 9))
        self.lbl_total_rows.pack(anchor='w', padx=10, pady=2)

        # 진행률 바
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.stats_frame, variable=self.progress_var,
                                           maximum=100, mode='determinate')
        self.progress_bar.pack(fill=tk.X, padx=10, pady=5)

        # 변수 목록 (체크박스 포함)
        list_frame = tk.LabelFrame(parent, text="변수 목록 (다중 선택 가능)",
                                  font=('맑은 고딕', 10, 'bold'),
                                  bg='white', fg=self.colors['text_dark'])
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 스크롤바가 있는 트리뷰
        tree_scroll = tk.Scrollbar(list_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_columns = ttk.Treeview(list_frame, columns=('Variable', 'Status', 'Unique'),
                                        show='tree headings', selectmode='extended',
                                        yscrollcommand=tree_scroll.set)

        self.tree_columns.heading('#0', text='✓')
        self.tree_columns.heading('Variable', text='변수명')
        self.tree_columns.heading('Status', text='상태')
        self.tree_columns.heading('Unique', text='고유값 수')

        self.tree_columns.column('#0', width=30, stretch=False)
        self.tree_columns.column('Variable', width=200)
        self.tree_columns.column('Status', width=80)
        self.tree_columns.column('Unique', width=80)

        self.tree_columns.pack(fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.tree_columns.yview)

        # 트리뷰 태그 스타일
        self.tree_columns.tag_configure('mapped', background='#C8E6C9')
        self.tree_columns.tag_configure('unmapped', background='#FFECB3')

        # 선택 버튼들
        select_frame = tk.Frame(parent, bg='white')
        select_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(select_frame, text="전체 선택", command=self.select_all_columns,
                 bg=self.colors['secondary'], fg='white', font=('맑은 고딕', 9),
                 relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)

        tk.Button(select_frame, text="선택 해제", command=self.deselect_all_columns,
                 bg=self.colors['warning'], fg='white', font=('맑은 고딕', 9),
                 relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)

        tk.Button(select_frame, text="매핑 안된 것만", command=self.select_unmapped,
                 bg=self.colors['primary'], fg='white', font=('맑은 고딕', 9),
                 relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=2)

        # 매핑 실행 버튼
        mapping_frame = tk.Frame(parent, bg='white')
        mapping_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(mapping_frame, text="🚀 선택한 변수 매핑 시작",
                 command=self.perform_mapping,
                 bg=self.colors['primary'], fg='white',
                 font=('맑은 고딕', 11, 'bold'),
                 relief=tk.FLAT, cursor='hand2', height=2).pack(fill=tk.X)

        tk.Button(mapping_frame, text="⚡ 빠른 매핑 (1→1, 2→2...)",
                 command=self.quick_mapping,
                 bg=self.colors['secondary'], fg='white',
                 font=('맑은 고딕', 10, 'bold'),
                 relief=tk.FLAT, cursor='hand2').pack(fill=tk.X, pady=(5, 0))

    def create_right_panel(self, parent):
        """우측 패널 생성 (미리보기 및 매핑 정보)"""
        # 탭 위젯
        notebook = ttk.Notebook(parent)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 탭 1: 데이터 미리보기
        preview_tab = tk.Frame(notebook, bg='white')
        notebook.add(preview_tab, text='📊 데이터 미리보기')

        preview_title = tk.Frame(preview_tab, bg=self.colors['primary'], height=35)
        preview_title.pack(fill=tk.X)
        preview_title.pack_propagate(False)

        tk.Label(preview_title, text="데이터 미리보기 (최대 100행)",
                bg=self.colors['primary'], fg='white',
                font=('맑은 고딕', 10, 'bold')).pack(pady=5)

        # 데이터 테이블
        preview_frame = tk.Frame(preview_tab)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scroll_x = tk.Scrollbar(preview_frame, orient=tk.HORIZONTAL)
        scroll_y = tk.Scrollbar(preview_frame, orient=tk.VERTICAL)

        self.tree_preview = ttk.Treeview(preview_frame,
                                        xscrollcommand=scroll_x.set,
                                        yscrollcommand=scroll_y.set)

        scroll_x.config(command=self.tree_preview.xview)
        scroll_y.config(command=self.tree_preview.yview)

        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_preview.pack(fill=tk.BOTH, expand=True)

        # 탭 2: 매핑 정보
        mapping_tab = tk.Frame(notebook, bg='white')
        notebook.add(mapping_tab, text='🔄 매핑 정보')

        mapping_title = tk.Frame(mapping_tab, bg=self.colors['secondary'], height=35)
        mapping_title.pack(fill=tk.X)
        mapping_title.pack_propagate(False)

        tk.Label(mapping_title, text="현재 매핑 정보",
                bg=self.colors['secondary'], fg='white',
                font=('맑은 고딕', 10, 'bold')).pack(pady=5)

        # 매핑 정보 테이블
        mapping_frame = tk.Frame(mapping_tab)
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        scroll_mapping = tk.Scrollbar(mapping_frame)
        scroll_mapping.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_mapping = ttk.Treeview(mapping_frame,
                                        columns=('Variable', 'Original', 'Mapped'),
                                        show='headings',
                                        yscrollcommand=scroll_mapping.set)

        self.tree_mapping.heading('Variable', text='변수명')
        self.tree_mapping.heading('Original', text='원본 값')
        self.tree_mapping.heading('Mapped', text='매핑된 값')

        self.tree_mapping.column('Variable', width=150)
        self.tree_mapping.column('Original', width=250)
        self.tree_mapping.column('Mapped', width=100)

        self.tree_mapping.pack(fill=tk.BOTH, expand=True)
        scroll_mapping.config(command=self.tree_mapping.yview)

        # 탭 3: 로그
        log_tab = tk.Frame(notebook, bg='white')
        notebook.add(log_tab, text='📝 작업 로그')

        log_scroll = tk.Scrollbar(log_tab)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_log = tk.Text(log_tab, wrap=tk.WORD, font=('맑은 고딕', 9),
                               yscrollcommand=log_scroll.set, bg='#FAFAFA')
        self.text_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        log_scroll.config(command=self.text_log.yview)

        # 로그 태그 설정
        self.text_log.tag_config('info', foreground='blue')
        self.text_log.tag_config('success', foreground='green', font=('맑은 고딕', 9, 'bold'))
        self.text_log.tag_config('warning', foreground='orange')
        self.text_log.tag_config('error', foreground='red', font=('맑은 고딕', 9, 'bold'))

    def create_statusbar(self, parent):
        """하단 상태바 생성"""
        statusbar = tk.Frame(parent, bg=self.colors['bg_dark'], height=30)
        statusbar.pack(fill=tk.X, side=tk.BOTTOM)
        statusbar.pack_propagate(False)

        self.lbl_status = tk.Label(statusbar, text="준비",
                                  bg=self.colors['bg_dark'], fg='white',
                                  font=('맑은 고딕', 9), anchor='w')
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        self.lbl_time = tk.Label(statusbar, text="",
                                bg=self.colors['bg_dark'], fg='white',
                                font=('맑은 고딕', 9))
        self.lbl_time.pack(side=tk.RIGHT, padx=10)
        self.update_time()

    def update_time(self):
        """시계 업데이트"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.lbl_time.config(text=current_time)
        self.root.after(1000, self.update_time)

    def add_hover_effect(self, button, original_color):
        """버튼 호버 효과"""
        def on_enter(e):
            button['bg'] = self.lighten_color(original_color)

        def on_leave(e):
            button['bg'] = original_color

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def lighten_color(self, color):
        """색상 밝게 만들기"""
        color_map = {
            self.colors['primary']: '#42A5F5',
            self.colors['secondary']: '#66BB6A',
            self.colors['warning']: '#FFA726',
            self.colors['danger']: '#EF5350',
            self.colors['success']: '#66BB6A'
        }
        return color_map.get(color, color)

    def drop_file(self, event):
        """드래그 앤 드롭으로 파일 업로드"""
        file_path = event.data.strip('{}')
        self.file_path = file_path
        self.load_file(file_path)

    def select_file(self):
        """파일 선택 대화상자"""
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.load_file(file_path)

    def load_file(self, file_path):
        """파일 로드"""
        try:
            self.log_message(f"파일 로드 중: {os.path.basename(file_path)}", 'info')
            self.update_status("파일 로딩 중...")

            self.df = pd.read_excel(file_path)
            self.original_df = self.df.copy()

            self.entry_file_path.delete(0, tk.END)
            self.entry_file_path.insert(0, file_path)

            # 통계 업데이트
            self.update_statistics()

            # 변수 목록 업데이트
            self.populate_column_list()

            # 미리보기 업데이트
            self.refresh_preview()

            self.log_message(f"✅ 파일 로드 완료: {len(self.df)} 행, {len(self.df.columns)} 변수", 'success')
            self.update_status(f"파일 로드 완료: {os.path.basename(file_path)}")

            messagebox.showinfo("성공", f"파일이 성공적으로 로드되었습니다!\n\n행: {len(self.df)}\n변수: {len(self.df.columns)}")

        except Exception as e:
            self.log_message(f"❌ 파일 로드 실패: {str(e)}", 'error')
            self.update_status("파일 로드 실패")
            messagebox.showerror("오류", f"파일을 로드할 수 없습니다:\n{str(e)}")

    def populate_column_list(self):
        """변수 목록 채우기"""
        self.tree_columns.delete(*self.tree_columns.get_children())

        if self.df is None:
            return

        for idx, col in enumerate(self.df.columns):
            # 고유값 개수 계산
            unique_count = len(self.get_unique_values(col))

            # 매핑 상태 확인
            is_mapped = col in self.user_defined_mappings
            status = "완료" if is_mapped else "대기"
            tag = 'mapped' if is_mapped else 'unmapped'

            self.tree_columns.insert('', 'end', iid=str(idx),
                                    text='☑' if is_mapped else '☐',
                                    values=(col, status, unique_count),
                                    tags=(tag,))

    def get_unique_values(self, col):
        """컬럼의 고유값 추출"""
        unique_values = set()

        for cell in self.df[col].dropna():
            if isinstance(cell, str):
                if cell.replace('.', '', 1).isdigit():
                    continue
                unique_values.update(cell.split("|"))
            elif isinstance(cell, (int, float)):
                continue

        return sorted(unique_values)

    def update_statistics(self):
        """통계 정보 업데이트"""
        if self.df is None:
            return

        total_vars = len(self.df.columns)
        mapped_vars = len(self.user_defined_mappings)
        total_rows = len(self.df)

        self.lbl_total_vars.config(text=f"총 변수: {total_vars}")
        self.lbl_mapped_vars.config(text=f"매핑 완료: {mapped_vars} / {total_vars}")
        self.lbl_total_rows.config(text=f"총 행: {total_rows:,}")

        # 진행률 계산
        if total_vars > 0:
            progress = (mapped_vars / total_vars) * 100
            self.progress_var.set(progress)

    def refresh_preview(self):
        """데이터 미리보기 갱신"""
        if self.df is None:
            return

        # 기존 데이터 삭제
        self.tree_preview.delete(*self.tree_preview.get_children())

        # 컬럼 설정
        self.tree_preview['columns'] = list(self.df.columns)
        self.tree_preview['show'] = 'headings'

        for col in self.df.columns:
            self.tree_preview.heading(col, text=col)
            self.tree_preview.column(col, width=120)

        # 데이터 삽입 (최대 100행)
        for idx, row in self.df.head(100).iterrows():
            values = [str(val)[:50] if pd.notna(val) else '' for val in row]
            self.tree_preview.insert('', 'end', values=values)

        self.log_message("🔄 데이터 미리보기 갱신", 'info')

    def filter_columns(self, *args):
        """검색어로 변수 필터링"""
        search_text = self.search_var.get().lower()

        self.tree_columns.delete(*self.tree_columns.get_children())

        if self.df is None:
            return

        for idx, col in enumerate(self.df.columns):
            if search_text in col.lower():
                unique_count = len(self.get_unique_values(col))
                is_mapped = col in self.user_defined_mappings
                status = "완료" if is_mapped else "대기"
                tag = 'mapped' if is_mapped else 'unmapped'

                self.tree_columns.insert('', 'end', iid=str(idx),
                                        text='☑' if is_mapped else '☐',
                                        values=(col, status, unique_count),
                                        tags=(tag,))

    def select_all_columns(self):
        """모든 변수 선택"""
        for item in self.tree_columns.get_children():
            self.tree_columns.selection_add(item)

    def deselect_all_columns(self):
        """모든 선택 해제"""
        self.tree_columns.selection_remove(*self.tree_columns.selection())

    def select_unmapped(self):
        """매핑 안된 변수만 선택"""
        self.deselect_all_columns()
        for item in self.tree_columns.get_children():
            if self.tree_columns.item(item)['values'][1] == "대기":
                self.tree_columns.selection_add(item)

    def perform_mapping(self):
        """매핑 수행 (향상된 사용자 인터페이스)"""
        if self.df is None:
            messagebox.showerror("오류", "먼저 엑셀 파일을 로드하세요!")
            return

        selection = self.tree_columns.selection()
        if not selection:
            messagebox.showerror("오류", "매핑할 변수를 선택하세요!")
            return

        selected_columns = [self.tree_columns.item(item)['values'][0] for item in selection]

        self.log_message(f"📝 매핑 시작: {len(selected_columns)}개 변수", 'info')
        self.update_status(f"매핑 중... ({len(selected_columns)}개 변수)")

        shared_mapping = {}  # 선택된 변수들 간 공유 매핑

        for col_idx, col in enumerate(selected_columns, 1):
            unique_values = self.get_unique_values(col)

            if not unique_values:
                continue

            self.log_message(f"\n[{col_idx}/{len(selected_columns)}] '{col}' 매핑 중... (고유값: {len(unique_values)}개)", 'info')

            mapping = {}

            for value in unique_values:
                # 이미 매핑된 값 확인
                if col in self.user_defined_mappings and value in self.user_defined_mappings[col]:
                    mapping[value] = self.user_defined_mappings[col][value]
                elif value in shared_mapping:
                    mapping[value] = shared_mapping[value]
                elif value in self.skip_values:
                    mapping[value] = value
                else:
                    # 매핑 입력 다이얼로그
                    user_input = self.show_mapping_dialog(col, value, col_idx, len(selected_columns))

                    if user_input is None:
                        self.log_message("⚠ 사용자가 매핑을 취소했습니다.", 'warning')
                        return

                    if user_input.lower() == 'p':
                        mapping[value] = value
                        self.skip_values.add(value)
                        self.log_message(f"  '{value}' → 패스 (원본 유지)", 'warning')
                    else:
                        try:
                            mapped_value = int(user_input)
                            mapping[value] = mapped_value
                            shared_mapping[value] = mapped_value
                            self.log_message(f"  '{value}' → {mapped_value}", 'success')
                        except ValueError:
                            messagebox.showerror("오류", "숫자를 입력하거나 'p'를 입력하세요!")
                            return

            # 매핑 적용
            self.df[col] = self.df[col].apply(
                lambda x: ",".join(map(str, [mapping[val] for val in x.split("|")]))
                if isinstance(x, str) and not x.replace('.', '', 1).isdigit() else x
            )

            self.user_defined_mappings[col] = mapping

        # UI 업데이트
        self.populate_column_list()
        self.update_statistics()
        self.refresh_preview()
        self.update_mapping_info()

        self.log_message(f"\n✅ 매핑 완료! {len(selected_columns)}개 변수가 성공적으로 매핑되었습니다.", 'success')
        self.update_status("매핑 완료")

        messagebox.showinfo("완료", f"{len(selected_columns)}개 변수의 매핑이 완료되었습니다!")

    def show_mapping_dialog(self, col, value, current, total):
        """향상된 매핑 입력 다이얼로그"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"매핑 입력 ({current}/{total})")
        dialog.geometry("500x300")
        dialog.resizable(False, False)
        dialog.grab_set()

        # 중앙 정렬
        dialog.transient(self.root)
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (300 // 2)
        dialog.geometry(f"+{x}+{y}")

        result = [None]

        # 헤더
        header = tk.Frame(dialog, bg=self.colors['primary'], height=60)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(header, text=f"변수: {col}", bg=self.colors['primary'],
                fg='white', font=('맑은 고딕', 11, 'bold')).pack(pady=5)
        tk.Label(header, text=f"진행: {current} / {total}", bg=self.colors['primary'],
                fg='white', font=('맑은 고딕', 9)).pack()

        # 본문
        body = tk.Frame(dialog, bg='white')
        body.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(body, text="원본 값:", bg='white',
                font=('맑은 고딕', 10, 'bold')).pack(anchor='w')
        tk.Label(body, text=value, bg=self.colors['bg_light'],
                font=('맑은 고딕', 12), relief=tk.RIDGE,
                padx=10, pady=10).pack(fill=tk.X, pady=(5, 15))

        tk.Label(body, text="매핑할 숫자를 입력하세요:", bg='white',
                font=('맑은 고딕', 10, 'bold')).pack(anchor='w')

        entry_var = tk.StringVar()
        entry = tk.Entry(body, textvariable=entry_var, font=('맑은 고딕', 14),
                        justify='center', relief=tk.SOLID, borderwidth=2)
        entry.pack(fill=tk.X, pady=5)
        entry.focus()

        tk.Label(body, text="(패스하려면 'p' 입력)", bg='white',
                font=('맑은 고딕', 8), fg='gray').pack()

        # 버튼
        button_frame = tk.Frame(dialog, bg='white')
        button_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

        def on_ok():
            result[0] = entry_var.get()
            dialog.destroy()

        def on_cancel():
            result[0] = None
            dialog.destroy()

        tk.Button(button_frame, text="✓ 확인", command=on_ok,
                 bg=self.colors['success'], fg='white',
                 font=('맑은 고딕', 10, 'bold'),
                 relief=tk.FLAT, cursor='hand2', width=15).pack(side=tk.LEFT, padx=5)

        tk.Button(button_frame, text="✗ 취소", command=on_cancel,
                 bg=self.colors['danger'], fg='white',
                 font=('맑은 고딕', 10, 'bold'),
                 relief=tk.FLAT, cursor='hand2', width=15).pack(side=tk.RIGHT, padx=5)

        # Enter 키로 확인
        entry.bind('<Return>', lambda e: on_ok())

        dialog.wait_window()
        return result[0]

    def quick_mapping(self):
        """빠른 매핑 (1→1, 2→2, 3→3...)"""
        if self.df is None:
            messagebox.showerror("오류", "먼저 파일을 로드하세요!")
            return

        selection = self.tree_columns.selection()
        if not selection:
            messagebox.showerror("오류", "변수를 선택하세요!")
            return

        selected_columns = [self.tree_columns.item(item)['values'][0] for item in selection]

        confirm = messagebox.askyesno("확인",
            f"{len(selected_columns)}개 변수에 자동 매핑을 적용하시겠습니까?\n\n"
            "1 → 1, 2 → 2, 3 → 3... 형식으로 매핑됩니다.")

        if not confirm:
            return

        for col in selected_columns:
            unique_values = self.get_unique_values(col)

            if not unique_values:
                continue

            mapping = {}
            for value in unique_values:
                # 숫자로 변환 가능하면 그대로 사용
                try:
                    mapping[value] = int(value)
                except:
                    mapping[value] = value  # 변환 불가능하면 원본 유지

            # 매핑 적용
            self.df[col] = self.df[col].apply(
                lambda x: ",".join(map(str, [mapping[val] for val in x.split("|")]))
                if isinstance(x, str) and not x.replace('.', '', 1).isdigit() else x
            )

            self.user_defined_mappings[col] = mapping

        self.populate_column_list()
        self.update_statistics()
        self.refresh_preview()
        self.update_mapping_info()

        self.log_message(f"⚡ 빠른 매핑 완료: {len(selected_columns)}개 변수", 'success')
        messagebox.showinfo("완료", f"{len(selected_columns)}개 변수의 빠른 매핑이 완료되었습니다!")

    def update_mapping_info(self):
        """매핑 정보 테이블 업데이트"""
        self.tree_mapping.delete(*self.tree_mapping.get_children())

        for col, mapping in self.user_defined_mappings.items():
            for original, mapped in mapping.items():
                self.tree_mapping.insert('', 'end',
                                        values=(col, original, mapped))

    def save_to_excel(self):
        """엑셀로 저장"""
        if self.df is None:
            messagebox.showerror("오류", "저장할 데이터가 없습니다!")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"매핑완료_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not save_path:
            return

        try:
            self.update_status("파일 저장 중...")

            # 매핑 정보 데이터프레임 생성
            mapping_data = []
            for col, mapping in self.user_defined_mappings.items():
                for original, mapped in mapping.items():
                    mapping_data.append({
                        '컬럼명': col,
                        '원본 값': original,
                        '매핑된 값': mapped
                    })

            mapping_df = pd.DataFrame(mapping_data)

            # 컬럼명 중복 제거 (시각적 효과)
            if not mapping_df.empty:
                mapping_df_display = mapping_df.copy()
                mapping_df_display['컬럼명'] = mapping_df_display['컬럼명'].mask(
                    mapping_df_display['컬럼명'].duplicated(), ''
                )
            else:
                mapping_df_display = mapping_df

            # 저장 디렉토리
            save_dir = os.path.dirname(save_path)
            mapping_path = os.path.join(save_dir,
                f"매핑정보_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

            # 파일 저장
            self.df.to_excel(save_path, index=False)
            if not mapping_df_display.empty:
                mapping_df_display.to_excel(mapping_path, index=False)

            self.log_message(f"💾 저장 완료:\n  - 데이터: {save_path}\n  - 매핑정보: {mapping_path}", 'success')
            self.update_status("저장 완료")

            result = messagebox.askyesno("저장 완료",
                f"파일이 성공적으로 저장되었습니다!\n\n"
                f"데이터: {os.path.basename(save_path)}\n"
                f"매핑정보: {os.path.basename(mapping_path)}\n\n"
                "저장 위치를 여시겠습니까?")

            if result:
                os.startfile(save_dir)

        except Exception as e:
            self.log_message(f"❌ 저장 실패: {str(e)}", 'error')
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{str(e)}")

    def save_template(self):
        """매핑 템플릿 저장"""
        if not self.user_defined_mappings:
            messagebox.showwarning("경고", "저장할 매핑 정보가 없습니다!")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=f"매핑템플릿_{datetime.now().strftime('%Y%m%d')}.json"
        )

        if not save_path:
            return

        try:
            template = {
                'mappings': self.user_defined_mappings,
                'skip_values': list(self.skip_values),
                'created_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }

            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, ensure_ascii=False, indent=2)

            self.log_message(f"📄 템플릿 저장 완료: {save_path}", 'success')
            messagebox.showinfo("성공", "매핑 템플릿이 저장되었습니다!")

        except Exception as e:
            messagebox.showerror("오류", f"템플릿 저장 실패:\n{str(e)}")

    def load_template(self):
        """매핑 템플릿 불러오기"""
        file_path = filedialog.askopenfilename(
            title="템플릿 선택",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                template = json.load(f)

            self.user_defined_mappings = template.get('mappings', {})
            self.skip_values = set(template.get('skip_values', []))

            self.populate_column_list()
            self.update_statistics()
            self.update_mapping_info()

            self.log_message(f"📂 템플릿 로드 완료: {file_path}", 'success')
            messagebox.showinfo("성공",
                f"템플릿이 로드되었습니다!\n\n"
                f"매핑 변수: {len(self.user_defined_mappings)}개\n"
                f"생성일: {template.get('created_date', '알 수 없음')}")

        except Exception as e:
            messagebox.showerror("오류", f"템플릿 로드 실패:\n{str(e)}")

    def undo_mapping(self):
        """매핑 되돌리기"""
        if self.original_df is None:
            messagebox.showwarning("경고", "되돌릴 데이터가 없습니다!")
            return

        confirm = messagebox.askyesno("확인",
            "모든 매핑을 취소하고 원본 데이터로 되돌리시겠습니까?")

        if confirm:
            self.df = self.original_df.copy()
            self.user_defined_mappings.clear()
            self.skip_values.clear()

            self.populate_column_list()
            self.update_statistics()
            self.refresh_preview()
            self.update_mapping_info()

            self.log_message("🔄 모든 매핑이 취소되었습니다.", 'warning')
            self.update_status("매핑 되돌리기 완료")
            messagebox.showinfo("완료", "원본 데이터로 되돌렸습니다!")

    def focus_search(self):
        """검색창 포커스"""
        # 검색 Entry 위젯을 찾아서 포커스
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Entry) and widget.cget('textvariable'):
                widget.focus()
                break

    def log_message(self, message, level='info'):
        """로그 메시지 추가"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        formatted_message = f"[{timestamp}] {message}\n"

        self.text_log.insert(tk.END, formatted_message, level)
        self.text_log.see(tk.END)

    def update_status(self, message):
        """상태바 업데이트"""
        self.lbl_status.config(text=message)
        self.root.update_idletasks()


def main():
    """메인 함수"""
    try:
        root = TkinterDnD.Tk()
        app = SurveyMappingApp(root)
        root.mainloop()
    except Exception as e:
        print(f"프로그램 실행 오류: {e}")
        traceback.print_exc()

        # TkinterDnD가 없는 경우 일반 Tkinter 사용
        if "TkinterDnD" in str(e):
            print("\n⚠ TkinterDnD2가 설치되지 않았습니다.")
            print("드래그 앤 드롭 기능 없이 실행합니다...\n")

            root = tk.Tk()
            app = SurveyMappingApp(root)
            root.mainloop()


if __name__ == "__main__":
    main()

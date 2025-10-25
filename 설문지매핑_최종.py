import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import threading
import platform
import subprocess

class SurveyMappingApp:
    """설문지 데이터 매핑 애플리케이션"""

    def __init__(self, root):
        self.root = root
        self.root.title("📊 설문지 데이터 매핑 도구")

        # 화면 크기 최적화
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = min(1200, int(screen_width * 0.85))
        window_height = min(800, int(screen_height * 0.85))

        # 중앙에 위치시키기
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(900, 600)  # 최소 크기 설정

        # 데이터 관련 변수
        self.df = None
        self.original_df = None
        self.file_path = None
        self.user_defined_mappings = {}
        self.skip_values = set()
        self.mapping_history = []

        # 시스템 테마 자동 감지
        self.current_theme = self.detect_system_theme()

        # UI 스타일 설정
        self.setup_styles()

        # 루트 윈도우 배경색 설정
        self.root.configure(bg=self.colors['bg'])

        # UI 구성
        self.create_widgets()

        # 키보드 단축키 바인딩
        self.setup_shortcuts()

        # 창 닫기 이벤트
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def detect_system_theme(self):
        """시스템 테마 자동 감지"""
        system = platform.system()

        try:
            if system == "Darwin":  # macOS
                # macOS에서 다크모드 확인
                result = subprocess.run(
                    ['defaults', 'read', '-g', 'AppleInterfaceStyle'],
                    capture_output=True,
                    text=True
                )
                return 'dark' if result.returncode == 0 and 'Dark' in result.stdout else 'light'

            elif system == "Windows":  # Windows
                try:
                    import winreg
                    registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
                    key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                    value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                    winreg.CloseKey(key)
                    return 'light' if value == 1 else 'dark'
                except:
                    return 'light'

            else:  # Linux 등 기타 OS
                # 환경 변수나 시간 기반으로 추측
                from datetime import datetime
                hour = datetime.now().hour
                # 저녁 7시부터 아침 7시까지는 다크모드
                return 'dark' if hour >= 19 or hour < 7 else 'light'

        except:
            return 'light'  # 기본값

    def setup_styles(self):
        """UI 스타일 설정 - 개선된 색상 대비"""
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # 개선된 테마 정의 (더 나은 색상 대비)
        self.themes = {
            'light': {
                'primary': '#1976D2',
                'success': '#388E3C',
                'warning': '#F57C00',
                'danger': '#D32F2F',
                'bg': '#FAFAFA',
                'card': '#FFFFFF',
                'text': '#000000',  # 진한 검정색으로 가독성 향상
                'text_secondary': '#424242',
                'border': '#E0E0E0',
                'input_bg': '#FFFFFF',
                'input_fg': '#000000',
                'select_bg': '#1976D2',
                'select_fg': '#FFFFFF',
                'button_fg': '#FFFFFF',
                'frame_bg': '#FFFFFF'
            },
            'dark': {
                'primary': '#5E92F3',  # 밝은 남색
                'success': '#66BB6A',
                'warning': '#FFA726',
                'danger': '#EF5350',
                'bg': '#0A0D10',  # 더 진한 남색 계열 배경 (거의 검정)
                'card': '#0D1015',  # 카드 배경 - 매우 어두운 남색
                'text': '#FFFFFF',  # 완전한 흰색으로 변경
                'text_secondary': '#E5E7EB',  # 밝은 회색으로 변경
                'border': '#374151',  # 테두리 - 진한 회색
                'input_bg': '#1F2937',  # 입력 필드 - 진한 회색으로 변경
                'input_fg': '#FFFFFF',  # 흰색으로 변경
                'select_bg': '#3B82F6',  # 선택 영역 - 밝은 파란색
                'select_fg': '#FFFFFF',
                'button_fg': '#000000',  # 버튼 텍스트 - 검정색
                'frame_bg': '#0A0D10'  # 프레임 배경
            }
        }

        self.colors = self.themes[self.current_theme]
        self.apply_theme()

    def apply_theme(self):
        """테마 적용"""
        colors = self.colors

        # 루트 윈도우 배경
        self.root.configure(bg=colors['bg'])

        # ttk 스타일 설정 - 테두리 개선
        self.style.configure('TFrame',
                           background=colors['bg'],
                           borderwidth=0,
                           relief='flat')

        # PanedWindow 스타일
        self.style.configure('TPanedwindow',
                           background=colors['bg'],
                           sashbackground=colors['bg'],
                           sashrelief='flat',
                           sashwidth=8)

        self.style.configure('TLabel',
                           background=colors['bg'],
                           foreground=colors['text'],
                           font=('Arial', 13))  # 글자 크기 증가

        # 패널 프레임 스타일 (더 어두운 배경)
        self.style.configure('Panel.TFrame',
                           background=colors['bg'],  # 메인 배경과 같은 어두운 색
                           borderwidth=0,
                           relief='flat')

        # LabelFrame 스타일 - 남색 카드 스타일
        self.style.configure('TLabelFrame',
                           background=colors['card'],
                           foreground=colors['text'],
                           borderwidth=0,
                           relief='flat',
                           bordercolor=colors['card'])

        # LabelFrame 내부 배경 설정
        self.style.configure('TLabelFrame.Label',
                           background=colors['card'],
                           foreground=colors['text'],
                           font=('Arial', 14, 'bold'))  # 라벨 글자 크기 증가

        # 버튼 스타일 - 더 현대적이고 가독성 있게
        self.style.configure('TButton',
                           background='#2563EB',  # 밝은 파란색으로 변경
                           foreground='white',
                           borderwidth=0,
                           relief='flat',
                           focuscolor='none',
                           padding=(12, 8),
                           font=('Arial', 14, 'bold'))  # 버튼 글자 크기 증가
        self.style.map('TButton',
                      background=[('active', '#3B82F6'),
                                ('pressed', '#1D4ED8'),
                                ('disabled', '#1E293B')],  # 어두운 회색 비활성화
                      relief=[('pressed', 'sunken')],
                      foreground=[('disabled', '#64748B')])  # 연한 회색 텍스트

        # 툴바 버튼 스타일 (더 눈에 띄고 명확하게)
        self.style.configure('Toolbar.TButton',
                           background='#3B82F6',  # 밝은 파란색
                           foreground='white',
                           borderwidth=0,
                           relief='flat',
                           padding=(15, 10),
                           font=('Arial', 14, 'bold'))
        self.style.map('Toolbar.TButton',
                      background=[('active', '#60A5FA'),
                                ('pressed', '#2563EB'),
                                ('disabled', '#334155')],  # 비활성화 시 어두운 회색
                      foreground=[('disabled', '#64748B')],  # 비활성화 텍스트
                      relief=[('pressed', 'sunken')])

        # Entry 스타일
        self.style.configure('TEntry',
                           fieldbackground=colors['input_bg'],
                           foreground=colors['input_fg'],
                           borderwidth=1,
                           relief='solid',
                           insertcolor=colors['text'])

        # Progressbar 스타일
        self.style.configure('TProgressbar',
                           background=colors['primary'],
                           borderwidth=0,
                           lightcolor=colors['primary'],
                           darkcolor=colors['primary'])

        # Scrollbar 스타일
        self.style.configure('TScrollbar',
                           background=colors['frame_bg'],
                           bordercolor=colors['border'],
                           arrowcolor=colors['text_secondary'],
                           troughcolor=colors['input_bg'])


    def update_widget_colors(self):
        """위젯 색상 업데이트"""
        colors = self.colors

        # Text 위젯 업데이트
        if hasattr(self, 'file_info_text'):
            self.file_info_text.config(bg=colors['input_bg'], fg=colors['input_fg'])
        if hasattr(self, 'stats_text'):
            self.stats_text.config(bg=colors['input_bg'], fg=colors['input_fg'])
        if hasattr(self, 'listbox_columns'):
            self.listbox_columns.config(bg=colors['input_bg'], fg=colors['input_fg'],
                                       selectbackground=colors['select_bg'],
                                       selectforeground=colors['select_fg'])


    def create_widgets(self):
        """UI 위젯 생성"""

        # 메인 컨테이너
        main_container = tk.Frame(self.root, bg=self.colors['bg'])
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # 상단 툴바
        self.create_toolbar(main_container)

        # 메인 패널 (PanedWindow로 크기 조절 가능)
        paned_window = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        paned_window.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        main_container.rowconfigure(1, weight=1)
        main_container.columnconfigure(0, weight=1)

        # 좌측 패널
        left_panel = tk.Frame(paned_window, bg=self.colors['bg'])
        paned_window.add(left_panel, weight=1)

        # 우측 패널
        right_panel = tk.Frame(paned_window, bg=self.colors['bg'])
        paned_window.add(right_panel, weight=1)

        # 좌측 패널 구성
        self.create_file_section(left_panel)
        self.create_column_section(left_panel)

        # 우측 패널 구성
        self.create_info_section(right_panel)
        self.create_result_section(right_panel)

        # 하단 상태바
        self.create_statusbar(main_container)

    def create_toolbar(self, parent):
        """상단 툴바 생성 - 심플하게"""
        toolbar = tk.Frame(parent, bg=self.colors['bg'])
        toolbar.grid(row=0, column=0, sticky=(tk.W, tk.E))

        # 주요 버튼만 포함 - 더 크고 잘 보이게
        btn_style = {'padx': 15, 'pady': 10}

        open_btn = tk.Button(toolbar, text="📂 파일 열기", command=self.select_file,
                            font=('Arial', 16, 'bold'),
                            bg='#3B82F6', fg='black',  # 검정색 텍스트
                            relief='raised', borderwidth=2,
                            activebackground='#60A5FA',
                            activeforeground='black',
                            cursor='hand2',
                            width=12, height=2)
        open_btn.pack(side=tk.LEFT, **btn_style)

        save_btn = tk.Button(toolbar, text="💾 저장", command=self.save_to_excel,
                            font=('Arial', 16, 'bold'),
                            bg='#10B981', fg='black',  # 검정색 텍스트
                            relief='raised', borderwidth=2,
                            activebackground='#34D399',
                            activeforeground='black',
                            cursor='hand2',
                            width=10, height=2)
        save_btn.pack(side=tk.LEFT, **btn_style)

        self.undo_button = tk.Button(toolbar, text="↩️ 되돌리기", command=self.undo_mapping,
                                     font=('Arial', 16, 'bold'),
                                     bg='#374151', fg='#9CA3AF',  # 비활성화 색상
                                     relief='raised', borderwidth=2,
                                     state='disabled',
                                     cursor='hand2',
                                     width=12, height=2)
        self.undo_button.pack(side=tk.LEFT, **btn_style)

    def create_file_section(self, parent):
        """파일 선택 섹션"""
        # tk.Frame으로 변경하여 배경색 적용
        file_frame = tk.Frame(parent, bg=self.colors['card'], highlightbackground=self.colors['border'], highlightthickness=1)
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        # 제목 라벨
        title_label = tk.Label(file_frame, text="📁 파일 선택", bg=self.colors['card'], fg=self.colors['text'],
                              font=('Arial', 12, 'bold'), anchor='w')
        title_label.pack(fill=tk.X, padx=15, pady=(10, 5))

        # 내부 프레임
        inner_frame = tk.Frame(file_frame, bg=self.colors['card'])
        inner_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        # 파일 경로 표시
        self.file_path_var = tk.StringVar(value="파일을 선택해주세요...")
        path_frame = tk.Frame(inner_frame, bg=self.colors['input_bg'], highlightbackground=self.colors['border'], highlightthickness=1)
        path_frame.pack(fill=tk.X, pady=(0, 10))
        path_label = tk.Label(path_frame, textvariable=self.file_path_var,
                             bg=self.colors['input_bg'],
                             fg=self.colors['input_fg'],
                             font=('Arial', 12),
                             pady=8, padx=8)
        path_label.pack(fill=tk.X)

        # 파일 정보
        self.file_info_text = tk.Text(inner_frame, height=3, width=40, wrap=tk.WORD,
                                     bg=self.colors['input_bg'], fg=self.colors['input_fg'],
                                     font=('Arial', 12),  # 글자 크기 증가
                                     relief='flat',
                                     borderwidth=1)
        self.file_info_text.pack(fill=tk.X)
        self.file_info_text.config(state='disabled')

    def create_column_section(self, parent):
        """컬럼 선택 섹션"""
        # tk.Frame으로 변경하여 배경색 적용
        column_frame = tk.Frame(parent, bg=self.colors['card'], highlightbackground=self.colors['border'], highlightthickness=1)
        column_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 제목 라벨
        title_label = tk.Label(column_frame, text="📋 변수 선택", bg=self.colors['card'], fg=self.colors['text'],
                              font=('Arial', 12, 'bold'), anchor='w')
        title_label.pack(fill=tk.X, padx=10, pady=(10, 5))

        # 내부 프레임
        inner_frame = tk.Frame(column_frame, bg=self.colors['card'])
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # 안내 텍스트
        guide_label = tk.Label(inner_frame, text="💡 Shift+클릭: 연속선택 | Ctrl+클릭: 개별선택",
                             bg=self.colors['card'], fg=self.colors['text'],
                             font=('Arial', 12))
        guide_label.pack(pady=(0, 10))

        # 컬럼 리스트 (스크롤바 포함) - 더 크게
        list_frame = tk.Frame(inner_frame, bg=self.colors['card'])
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 스크롤바를 위한 tk 사용 (더 나은 색상 제어)
        scrollbar = tk.Scrollbar(list_frame, bg=self.colors['card'],
                                activebackground=self.colors['primary'],
                                troughcolor=self.colors['input_bg'])
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox_columns = tk.Listbox(list_frame, selectmode=tk.EXTENDED,
                                         yscrollcommand=scrollbar.set,
                                         bg=self.colors['input_bg'],
                                         fg=self.colors['input_fg'],
                                         selectbackground=self.colors['primary'],
                                         selectforeground='white',
                                         font=('Arial', 14),  # 글자 크기 더 증가
                                         height=15,
                                         relief='flat',
                                         borderwidth=0,
                                         highlightbackground=self.colors['card'],
                                         highlightcolor=self.colors['primary'],
                                         highlightthickness=0)
        self.listbox_columns.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox_columns.yview)

        # 매핑하기 버튼 - 크고 눈에 띄게
        button_frame = tk.Frame(inner_frame, bg=self.colors['card'])
        button_frame.pack(fill=tk.X, pady=(15, 0))

        # 다크모드 체크
        text_color = 'black' if self.current_theme == 'dark' else 'white'

        mapping_button = tk.Button(button_frame,
                                 text="📝 매핑하기",
                                 command=self.perform_batch_mapping,
                                 font=('Arial', 20, 'bold'),  # 더 큰 글자
                                 bg='#8B5CF6',  # 보라색으로 변경하여 눈에 띄게
                                 fg=text_color,  # 다크모드일 때 검정색
                                 height=3,  # 높이 증가
                                 cursor='hand2',
                                 relief='raised',
                                 borderwidth=3,
                                 activebackground='#A78BFA',
                                 activeforeground=text_color,
                                 disabledforeground='#9CA3AF',  # 비활성화 텍스트 색상
                                 highlightthickness=0)
        mapping_button.pack(fill=tk.X, padx=20)

    def create_info_section(self, parent):
        """정보 표시 섹션"""
        # tk.Frame으로 변경하여 배경색 적용
        info_frame = tk.Frame(parent, bg=self.colors['card'], highlightbackground=self.colors['border'], highlightthickness=1)
        info_frame.pack(fill=tk.X, padx=5, pady=5)

        # 제목 라벨
        title_label = tk.Label(info_frame, text="ℹ️ 정보", bg=self.colors['card'], fg=self.colors['text'],
                              font=('Arial', 12, 'bold'), anchor='w')
        title_label.pack(fill=tk.X, padx=10, pady=(10, 5))

        # 내부 프레임
        inner_frame = tk.Frame(info_frame, bg=self.colors['card'])
        inner_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.stats_text = ScrolledText(inner_frame, height=8, width=40, wrap=tk.WORD,
                                      bg=self.colors['input_bg'], fg=self.colors['input_fg'],
                                      font=('Arial', 13),  # 글자 크기 증가
                                      relief='flat',
                                      borderwidth=0)
        self.stats_text.pack(fill=tk.BOTH, expand=True)
        self.stats_text.config(state='disabled')

    def create_result_section(self, parent):
        """결과 표시 섹션"""
        # tk.Frame으로 변경하여 배경색 적용
        result_frame = tk.Frame(parent, bg=self.colors['card'], highlightbackground=self.colors['border'], highlightthickness=1)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 제목 라벨
        title_label = tk.Label(result_frame, text="📝 매핑 결과", bg=self.colors['card'], fg=self.colors['text'],
                              font=('Arial', 12, 'bold'), anchor='w')
        title_label.pack(fill=tk.X, padx=10, pady=(10, 5))

        # 내부 프레임
        inner_frame = tk.Frame(result_frame, bg=self.colors['card'])
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # 결과 텍스트
        self.result_text = ScrolledText(inner_frame, height=10, width=40, wrap=tk.WORD,
                                       bg=self.colors['input_bg'], fg=self.colors['input_fg'],
                                       font=('Arial', 13),  # 글자 크기 증가
                                       relief='flat',
                                       borderwidth=0)
        self.result_text.pack(fill=tk.BOTH, expand=True)
        self.result_text.config(state='disabled')

    def create_statusbar(self, parent):
        """하단 상태바"""
        statusbar = tk.Frame(parent, bg=self.colors['bg'])
        statusbar.grid(row=2, column=0, sticky=(tk.W, tk.E))

        self.status_var = tk.StringVar(value="준비됨")
        ttk.Label(statusbar, textvariable=self.status_var, relief='sunken').pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(statusbar, variable=self.progress_var,
                                           length=200, mode='determinate')
        self.progress_bar.pack(side=tk.RIGHT, padx=(10, 0))

    def setup_shortcuts(self):
        """키보드 단축키 설정"""
        self.root.bind('<Control-o>', lambda _: self.select_file())
        self.root.bind('<Control-s>', lambda _: self.save_to_excel())
        self.root.bind('<Control-z>', lambda _: self.undo_mapping())

    def select_file(self):
        """엑셀 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if not file_path:
            return

        self.file_path = file_path
        self.file_path_var.set(os.path.basename(file_path))

        self.status_var.set("파일 로딩 중...")
        self.progress_bar.start()

        thread = threading.Thread(target=self.load_file_async, args=(file_path,))
        thread.start()

    def load_file_async(self, file_path):
        """비동기 파일 로드"""
        try:
            self.df = pd.read_excel(file_path)
            self.original_df = self.df.copy()
            self.root.after(0, self.update_after_load, True)
        except Exception as e:
            self.root.after(0, self.update_after_load, False, str(e))

    def update_after_load(self, success, error_msg=None):
        """파일 로드 후 UI 업데이트"""
        self.progress_bar.stop()

        if success:
            # 컬럼 리스트 업데이트
            self.listbox_columns.delete(0, tk.END)
            for col in self.df.columns:
                self.listbox_columns.insert(tk.END, col)

            # 파일 정보 표시
            self.update_file_info()

            # 통계 정보 표시
            self.update_stats_info()

            self.status_var.set("파일 로드 완료")
            messagebox.showinfo("성공", "엑셀 파일이 성공적으로 로드되었습니다!")
        else:
            self.status_var.set("파일 로드 실패")
            messagebox.showerror("오류", f"파일을 열 수 없습니다: {error_msg}")

    def update_file_info(self):
        """파일 정보 업데이트"""
        if self.df is None:
            return

        info = f"파일: {os.path.basename(self.file_path)}\n"
        info += f"행: {len(self.df):,} | 열: {len(self.df.columns):,}\n"
        info += f"크기: {self.df.memory_usage().sum() / 1024**2:.2f} MB"

        self.file_info_text.config(state='normal')
        self.file_info_text.delete(1.0, tk.END)
        self.file_info_text.insert(1.0, info)
        self.file_info_text.config(state='disabled')

    def update_stats_info(self):
        """통계 정보 업데이트"""
        if self.df is None:
            return

        info = "📊 데이터 요약\n\n"

        # 데이터 타입 요약
        dtype_counts = self.df.dtypes.value_counts()
        for dtype, count in dtype_counts.items():
            info += f"{dtype}: {count}개\n"

        # 결측치 현황
        null_total = self.df.isnull().sum().sum()
        if null_total > 0:
            info += f"\n결측치: {null_total:,}개"

        self.stats_text.config(state='normal')
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(1.0, info)
        self.stats_text.config(state='disabled')


    def perform_batch_mapping(self):
        """일괄 매핑 수행"""
        if self.df is None:
            messagebox.showerror("오류", "먼저 엑셀 파일을 선택하세요!")
            return

        selected_columns = [self.listbox_columns.get(idx)
                          for idx in self.listbox_columns.curselection()]
        if not selected_columns:
            messagebox.showerror("오류", "매핑할 변수를 선택하세요!")
            return

        # 모든 고유값 수집
        all_unique_values = set()
        column_values = {}

        for col in selected_columns:
            unique_values = self.extract_unique_values(col)
            column_values[col] = unique_values
            all_unique_values.update(unique_values)

        if not all_unique_values:
            messagebox.showinfo("정보", "매핑할 값이 없습니다.")
            return

        # 일괄 매핑 다이얼로그 표시
        dialog = BatchMappingDialog(self.root, sorted(all_unique_values), self.skip_values, self.colors)
        self.root.wait_window(dialog.dialog)

        if dialog.result is None:
            self.status_var.set("매핑 취소됨")
            return

        # 매핑 전 백업
        before_mapping = self.df.copy()

        try:
            batch_mapping = dialog.result

            # 각 컬럼에 매핑 적용
            for idx, col in enumerate(selected_columns):
                progress = (idx / len(selected_columns)) * 100
                self.progress_var.set(progress)
                self.status_var.set(f"매핑 중... ({idx+1}/{len(selected_columns)})")
                self.root.update()

                # 컬럼별 매핑 생성
                col_mapping = {}
                for value in column_values[col]:
                    if value in batch_mapping:
                        col_mapping[value] = batch_mapping[value]
                    else:
                        col_mapping[value] = value

                # 매핑 적용
                self.apply_mapping(col, col_mapping)
                self.user_defined_mappings[col] = col_mapping

            # 매핑 히스토리 저장
            self.mapping_history.append({
                'before': before_mapping,
                'after': self.df.copy(),
                'mappings': self.user_defined_mappings.copy()
            })

            # 되돌리기 버튼 활성화
            self.undo_button.config(state='normal', bg='#EF4444', fg=getattr(self, 'undo_button_text_color', 'black' if self.current_theme == 'dark' else 'white'))  # 활성화

            # 결과 표시
            self.update_result_display()

            self.progress_var.set(100)
            self.status_var.set("매핑 완료")
            messagebox.showinfo("완료", "매핑이 완료되었습니다!")

        except Exception as e:
            self.df = before_mapping
            self.status_var.set("매핑 실패")
            self.progress_var.set(0)
            messagebox.showerror("오류", f"매핑 중 오류가 발생했습니다:\n{str(e)}")

    def extract_unique_values(self, column):
        """컬럼에서 고유값 추출"""
        unique_values = set()

        for cell in self.df[column].dropna():
            if isinstance(cell, str):
                if cell.replace('.', '', 1).isdigit():
                    continue
                unique_values.update(cell.split("|"))
            elif isinstance(cell, (int, float)):
                continue

        return sorted(unique_values)

    def apply_mapping(self, column, mapping):
        """매핑 적용"""
        def map_value(x):
            if isinstance(x, str) and not x.replace('.', '', 1).isdigit():
                mapped_values = [str(mapping.get(val, val)) for val in x.split("|")]
                return ",".join(mapped_values)
            return x

        self.df[column] = self.df[column].apply(map_value)

    def update_result_display(self):
        """매핑 결과 표시"""
        if not self.user_defined_mappings:
            return

        result = "📊 매핑 완료\n\n"
        for col, mappings in self.user_defined_mappings.items():
            result += f"[{col}]\n"
            for original, mapped in list(mappings.items())[:5]:  # 처음 5개만 표시
                result += f"  {original} → {mapped}\n"
            if len(mappings) > 5:
                result += f"  ... 외 {len(mappings)-5}개\n"
            result += "\n"

        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, result)
        self.result_text.config(state='disabled')

    def undo_mapping(self):
        """매핑 되돌리기 - 완전히 처음 상태로 초기화"""
        if not self.mapping_history:
            return

        # 매핑 히스토리 모두 삭제
        self.mapping_history.clear()

        # 원본 데이터로 복구
        self.df = self.original_df.copy() if self.original_df is not None else None

        # 매핑 설정 완전 초기화
        self.user_defined_mappings = {}
        self.skip_values = set()

        # 되돌리기 버튼 비활성화
        self.undo_button.config(state='disabled', bg='#374151', fg='#9CA3AF')

        self.status_var.set("원본 데이터로 복구 완료")

        # 결과 창 초기화 - 아무것도 표시하지 않음
        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state='disabled')

        # 통계 업데이트
        self.update_stats()

        messagebox.showinfo("완료", "원본 데이터로 복구되었습니다.\n새로 매핑을 시작하세요.")

    def save_to_excel(self):
        """엑셀로 저장"""
        if self.df is None:
            messagebox.showerror("오류", "저장할 데이터가 없습니다!")
            return

        # 원본 파일명에서 확장자 제거하고 "_매핑완료" 추가
        if self.file_path:
            base_name = os.path.splitext(os.path.basename(self.file_path))[0]
            default_name = f"{base_name}_매핑완료.xlsx"
        else:
            default_name = f"매핑완료_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=default_name,
            initialdir=os.path.dirname(self.file_path) if self.file_path else None
        )

        if not save_path:
            return

        try:
            self.status_var.set("저장 중...")

            # 메인 데이터 저장
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                self.df.to_excel(writer, sheet_name='매핑된 데이터', index=False)

                # 매핑 정보 저장
                if self.user_defined_mappings:
                    mapping_data = []
                    for col, mappings in self.user_defined_mappings.items():
                        for original, mapped in mappings.items():
                            mapping_data.append({
                                '컬럼명': col,
                                '원본 값': original,
                                '매핑된 값': mapped
                            })

                    if mapping_data:
                        mapping_df = pd.DataFrame(mapping_data)
                        mapping_df.to_excel(writer, sheet_name='매핑 정보', index=False)

            self.status_var.set("저장 완료")
            messagebox.showinfo("성공", f"파일이 저장되었습니다:\n{os.path.basename(save_path)}")

        except Exception as e:
            self.status_var.set("저장 실패")
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다:\n{str(e)}")

    def on_closing(self):
        """프로그램 종료 시"""
        if self.df is not None and self.user_defined_mappings:
            result = messagebox.askyesnocancel(
                "종료",
                "저장하지 않은 변경사항이 있습니다.\n저장하시겠습니까?"
            )

            if result is True:
                self.save_to_excel()
            elif result is None:
                return

        self.root.destroy()


class BatchMappingDialog:
    """일괄 매핑 입력 다이얼로그"""

    def __init__(self, parent, unique_values, skip_values, colors):
        self.parent = parent  # parent 저장
        self.result = None
        self.unique_values = unique_values
        self.skip_values = skip_values
        self.entry_widgets = {}
        self.colors = colors

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("📝 일괄 매핑 설정")

        # 화면 크기에 맞춰 조정 - 적당한 크기
        screen_width = self.dialog.winfo_screenwidth()
        screen_height = self.dialog.winfo_screenheight()
        dialog_width = min(1400, int(screen_width * 0.8))  # 적당한 너비
        dialog_height = min(1000, int(screen_height * 0.85))  # 적당한 높이

        x = (screen_width - dialog_width) // 2
        y = (screen_height - dialog_height) // 2

        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        self.dialog.configure(bg=colors['bg'])

        # 중앙 배치
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 다이얼로그 전체에서 엔터키 바인딩
        self.dialog.bind('<Return>', lambda e: self.ok_clicked())
        # 모든 하위 위젯에서도 엔터키가 작동하도록 설정
        self.dialog.focus_force()

        # 메인 프레임
        main_frame = tk.Frame(self.dialog, bg=colors['bg'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 상단 설명
        info_frame = tk.Frame(main_frame, bg=colors['bg'])
        info_frame.pack(fill=tk.X, pady=(0, 10))

        title_label = tk.Label(info_frame, text="📌 일괄 매핑 설정", font=('Arial', 24, 'bold'),  # 적당한 크기
                              bg=colors['bg'], fg=colors['text'])
        title_label.pack(anchor='w')

        desc_label = tk.Label(info_frame, text="각 값에 대한 매핑 숫자를 입력하세요. 빈 칸이나 'p'는 변환하지 않습니다.",
                             font=('Arial', 16), bg=colors['bg'], fg=colors['text_secondary'])  # 적당한 크기
        desc_label.pack(anchor='w', pady=(5, 0))

        # 빠른 설정 버튼
        quick_frame = tk.Frame(main_frame, bg=colors['bg'])
        quick_frame.pack(fill=tk.X, pady=(10, 20))  # 상하 간격 증가

        auto_btn = tk.Button(quick_frame, text="🔢 자동 번호", command=self.auto_number,
                            font=('Arial', 18, 'bold'),  # 적당한 크기
                            bg='#60A5FA', fg='black',  # 밝은 파란색 + 검은 텍스트
                            relief='raised', borderwidth=3,
                            activebackground='#93C5FD',
                            cursor='hand2', padx=30, pady=12,  # 적당한 패딩
                            height=2, width=12)  # 적당한 크기
        auto_btn.pack(side=tk.LEFT, padx=15)

        clear_btn = tk.Button(quick_frame, text="🔄 초기화", command=self.clear_all,
                             font=('Arial', 18, 'bold'),  # 적당한 크기
                             bg='#F87171', fg='black',  # 밝은 빨간색 + 검은 텍스트
                             relief='raised', borderwidth=3,
                             activebackground='#FCA5A5',
                             cursor='hand2', padx=30, pady=12,  # 적당한 패딩
                             height=2, width=10)  # 적당한 크기
        clear_btn.pack(side=tk.LEFT, padx=10)

        # 스크롤 가능한 매핑 입력 영역
        canvas_frame = tk.Frame(main_frame, bg=colors['card'])
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        # 캔버스와 스크롤바
        canvas = tk.Canvas(canvas_frame, bg=colors['card'], highlightthickness=0)
        scrollbar = tk.Scrollbar(canvas_frame, orient='vertical', command=canvas.yview,
                                bg='#374151', activebackground='#6B7280',
                                troughcolor='#1F2937')
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 내부 프레임
        self.inner_frame = tk.Frame(canvas, bg=colors['card'])
        canvas_window = canvas.create_window((0, 0), window=self.inner_frame, anchor='nw')

        # 매핑 입력 필드 생성
        self.create_mapping_fields()

        # 캔버스 크기 조정
        def configure_canvas(_=None):
            canvas.configure(scrollregion=canvas.bbox('all'))
            canvas_width = canvas.winfo_width()
            if self.inner_frame.winfo_reqwidth() < canvas_width:
                canvas.itemconfig(canvas_window, width=canvas_width)

        self.inner_frame.bind('<Configure>', configure_canvas)
        canvas.bind('<Configure>', lambda _: configure_canvas())

        # 하단 버튼
        button_frame = tk.Frame(main_frame, bg=colors['bg'])
        button_frame.pack(fill=tk.X, pady=(10, 0))

        ok_btn = tk.Button(button_frame, text="✅ 적용하기", command=self.ok_clicked,
                          font=('Arial', 18, 'bold'),  # 적당한 크기
                          bg='#34D399', fg='black',  # 밝은 초록색 + 검은 텍스트
                          relief='raised', borderwidth=3,
                          activebackground='#6EE7B7',
                          cursor='hand2', padx=30, pady=15,  # 적당한 패딩
                          height=2,  # 적당한 높이
                          width=12)  # 적당한 너비
        ok_btn.pack(side=tk.RIGHT, padx=15)

        cancel_btn = tk.Button(button_frame, text="❌ 취소", command=self.cancel_clicked,
                              font=('Arial', 18, 'bold'),  # 적당한 크기
                              bg='#FB923C', fg='black',  # 밝은 주황색 + 검은 텍스트
                              relief='raised', borderwidth=3,
                              activebackground='#FDBA74',
                              cursor='hand2', padx=30, pady=15,  # 적당한 패딩
                              height=2,  # 적당한 높이
                              width=10)  # 적당한 너비
        cancel_btn.pack(side=tk.RIGHT, padx=10)

    def create_mapping_fields(self):
        """매핑 입력 필드 생성"""
        colors = self.colors  # self.colors 사용
        for idx, value in enumerate(self.unique_values):
            row_frame = tk.Frame(self.inner_frame, bg=colors['card'])
            row_frame.pack(fill=tk.X, pady=8, padx=20)  # 적당한 간격

            # 원본 값 표시
            value_label = tk.Label(row_frame, text=str(value)[:60], width=45,  # 적당한 크기
                                  bg=colors['card'], fg=colors['text'],
                                  font=('Arial', 16, 'bold'), anchor='w')  # 적당한 크기
            value_label.pack(side=tk.LEFT, padx=15)  # 적당한 패딩

            arrow_label = tk.Label(row_frame, text="➡️",  # 화살표 이모지
                                  bg=colors['card'], fg=colors['text_secondary'],
                                  font=('Arial', 18, 'bold'))  # 적당한 크기
            arrow_label.pack(side=tk.LEFT, padx=15)

            # 매핑 값 입력
            entry = tk.Entry(row_frame, width=20,  # 적당한 너비
                            bg=colors['input_bg'], fg=colors['input_fg'],
                            font=('Arial', 16, 'bold'),  # 적당한 크기
                            insertbackground=colors['text'],
                            relief='solid', borderwidth=2,  # 적당한 테두리
                            highlightbackground='#374151',
                            highlightcolor='#3B82F6',  # 포커스시 파란색
                            highlightthickness=2)
            entry.pack(side=tk.LEFT, padx=10, ipady=5)  # 적당한 높이

            if value in self.skip_values:
                entry.insert(0, "p")

            self.entry_widgets[value] = entry

            # 엔터키 바인딩 - 모든 필드에서 엔터키를 누르면 적용
            entry.bind('<Return>', lambda e: self.ok_clicked())

    def auto_number(self):
        """자동 번호 매기기"""
        for idx, entry in enumerate(self.entry_widgets.values(), start=1):
            entry.delete(0, tk.END)
            entry.insert(0, str(idx))

    def clear_all(self):
        """모든 입력 초기화"""
        for entry in self.entry_widgets.values():
            entry.delete(0, tk.END)

    def ok_clicked(self):
        """적용 버튼"""
        self.result = {}

        for value, entry in self.entry_widgets.items():
            input_value = entry.get().strip()

            if not input_value or input_value.lower() == 'p':
                self.result[value] = value
            else:
                try:
                    self.result[value] = int(input_value)
                except ValueError:
                    messagebox.showerror("오류", f"'{value}'에 대한 값 '{input_value}'는 올바른 숫자가 아닙니다!")
                    return

        self.dialog.destroy()

    def cancel_clicked(self):
        """취소 버튼"""
        self.result = None
        self.dialog.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = SurveyMappingApp(root)
    root.mainloop()
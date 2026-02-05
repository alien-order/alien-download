import pandas as pd
import requests
import os
import time
import threading
from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.ttk import Progressbar


class FileDownloaderBase:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("대용량 파일 다운로더")
        self.root.geometry("900x700")
        
        # 스타일 설정
        self.setup_styles()
        
        self.is_paused = False
        self.is_stopped = False
        self.current_thread = None
        
        self.total_count = 0
        self.success_count = 0
        self.fail_count = 0
        self.current_index = 0
        
        self.result_data = []
        self.excel_file_path = ""
        self.start_row = 0
        self.start_col = 0
        
        self.setup_ui()
    
    def setup_styles(self):
        """UI 스타일 설정"""
        self.root.configure(bg='#f0f0f0')
        
        style = ttk.Style()
        style.theme_use('clam')
        
        # 버튼 스타일
        style.configure('Action.TButton', 
                       font=('Arial', 10, 'bold'),
                       padding=(10, 5))
        
        # 라벨프레임 스타일
        style.configure('Card.TLabelframe',
                       borderwidth=1,
                       relief='solid')
        
        # 진행률 바 스타일
        style.configure('Success.Horizontal.TProgressbar',
                       background='#4CAF50',
                       troughcolor='#E0E0E0',
                       borderwidth=1,
                       lightcolor='#4CAF50',
                       darkcolor='#4CAF50')
    
    def parse_cell_reference(self, cell_ref):
        """A2 형식의 셀 참조를 행과 열 인덱스로 변환"""
        cell_ref = cell_ref.upper().strip()
        
        col_str = ""
        row_str = ""
        
        for char in cell_ref:
            if char.isalpha():
                col_str += char
            elif char.isdigit():
                row_str += char
        
        if not col_str or not row_str:
            raise ValueError("올바른 셀 참조가 아닙니다 (예: A2, B10)")
        
        # 컬럼을 숫자로 변환 (A=0, B=1, ...)
        col_index = 0
        for char in col_str:
            col_index = col_index * 26 + (ord(char) - ord('A') + 1)
        col_index -= 1  # 0-based
        
        # 행을 숫자로 변환 (1-based에서 0-based로)
        row_index = int(row_str) - 1
        
        return row_index, col_index
    
    def setup_ui(self):
        # 메인 프레임 (스크롤 가능)
        canvas = tk.Canvas(self.root, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        main_frame = ttk.Frame(scrollable_frame, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # 파일 설정 프레임
        file_frame = ttk.LabelFrame(main_frame, text="📋 파일 설정", 
                                   style='Card.TLabelframe', padding="15")
        file_frame.pack(fill="x", pady=(0, 15))
        
        # 엑셀 파일 선택
        file_row = ttk.Frame(file_frame)
        file_row.pack(fill="x", pady=5)
        ttk.Label(file_row, text="엑셀 파일:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 10))
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_row, textvariable=self.file_path_var, font=('Arial', 9))
        file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(file_row, text="📂 찾기", command=self.select_excel_file, 
                  style='Action.TButton').pack(side="right")
        
        # 시작 셀 설정
        cell_row = ttk.Frame(file_frame)
        cell_row.pack(fill="x", pady=5)
        ttk.Label(cell_row, text="시작 셀:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 10))
        self.start_cell_var = tk.StringVar(value="A2")
        cell_entry = ttk.Entry(cell_row, textvariable=self.start_cell_var, width=10, font=('Arial', 9))
        cell_entry.pack(side="left", padx=(0, 10))
        ttk.Label(cell_row, text="(예: A2, B1, C5)", foreground="gray").pack(side="left")
        
        # 다운로드 폴더 설정
        folder_row = ttk.Frame(file_frame)
        folder_row.pack(fill="x", pady=5)
        ttk.Label(folder_row, text="저장 폴더:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 10))
        self.download_folder_var = tk.StringVar(value="./downloads")
        folder_entry = ttk.Entry(folder_row, textvariable=self.download_folder_var, font=('Arial', 9))
        folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Button(folder_row, text="📁 선택", command=self.select_download_folder,
                  style='Action.TButton').pack(side="right")
        
        # 딜레이 설정
        delay_row = ttk.Frame(file_frame)
        delay_row.pack(fill="x", pady=5)
        ttk.Label(delay_row, text="요청 딜레이(ms):", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 10))
        self.delay_var = tk.StringVar(value="100")
        ttk.Entry(delay_row, textvariable=self.delay_var, width=8, font=('Arial', 9)).pack(side="left", padx=(0, 10))
        ttk.Label(delay_row, text="(서버 부하 방지용)", foreground="gray").pack(side="left")
        
        # 이미지 없음 URL 설정
        noimage_row = ttk.Frame(file_frame)
        noimage_row.pack(fill="x", pady=5)
        ttk.Label(noimage_row, text="이미지 없음 URL:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 10))
        self.no_image_url_var = tk.StringVar(value="no_image")
        ttk.Entry(noimage_row, textvariable=self.no_image_url_var, width=30, font=('Arial', 9)).pack(side="left", padx=(0, 10))
        ttk.Label(noimage_row, text="(이 URL이면 다운로드 건너뜀)", foreground="gray").pack(side="left")
        
        # 제어 버튼 프레임
        control_frame = ttk.LabelFrame(main_frame, text="🎮 제어", 
                                     style='Card.TLabelframe', padding="15")
        control_frame.pack(fill="x", pady=(0, 15))
        
        button_container = ttk.Frame(control_frame)
        button_container.pack()
        
        self.start_btn = ttk.Button(button_container, text="▶️ 시작", command=self.start_download,
                                   style='Action.TButton', width=12)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.pause_btn = ttk.Button(button_container, text="⏸️ 일시정지", command=self.pause_download, 
                                   state="disabled", style='Action.TButton', width=12)
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        
        self.resume_btn = ttk.Button(button_container, text="▶️ 재개", command=self.resume_download, 
                                    state="disabled", style='Action.TButton', width=12)
        self.resume_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(button_container, text="⏹️ 중지", command=self.stop_download, 
                                  state="disabled", style='Action.TButton', width=12)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # 진행 상태 프레임
        progress_frame = ttk.LabelFrame(main_frame, text="📊 진행 상태", 
                                      style='Card.TLabelframe', padding="15")
        progress_frame.pack(fill="x", pady=(0, 15))
        
        # 상태 정보
        self.status_label = ttk.Label(progress_frame, text="🔘 준비됨", 
                                     font=('Arial', 11, 'bold'))
        self.status_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.detail_label = ttk.Label(progress_frame, text="", 
                                     font=('Arial', 9), foreground="#666666")
        self.detail_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 진행률 바
        progress_container = ttk.Frame(progress_frame)
        progress_container.pack(fill="x")
        
        ttk.Label(progress_container, text="전체 진행률:", font=('Arial', 9, 'bold')).pack(anchor=tk.W)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = Progressbar(progress_container, variable=self.progress_var, 
                                       style='Success.Horizontal.TProgressbar')
        self.progress_bar.pack(fill=tk.X, pady=(5, 0), ipady=5)
        
        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="📝 실행 로그", 
                                 style='Card.TLabelframe', padding="15")
        log_frame.pack(fill="both", expand=True)
        
        # 스크롤 가능한 텍스트 영역
        log_container = ttk.Frame(log_frame)
        log_container.pack(fill="both", expand=True)
        
        self.log_text = tk.Text(log_container, height=15, wrap=tk.WORD, 
                               font=('Consolas', 9), bg='#f8f8f8', 
                               selectbackground='#e3f2fd')
        log_scrollbar = ttk.Scrollbar(log_container, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.excel_file_path = file_path
    
    def select_download_folder(self):
        folder_path = filedialog.askdirectory(title="다운로드 폴더 선택")
        if folder_path:
            self.download_folder_var.set(folder_path)
    
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self):
        if self.total_count > 0:
            progress = ((self.success_count + self.fail_count) / self.total_count) * 100
            self.progress_var.set(progress)
            
            if self.is_paused:
                status_text = f"⏸️ 일시정지 - 진행중... ({self.current_index}/{self.total_count})"
            else:
                status_text = f"🚀 진행중... ({self.current_index}/{self.total_count})"
            
            detail_text = f"✅ 완료({self.success_count}/{self.total_count}), ❌ 실패({self.fail_count}/{self.total_count})"
            
            self.status_label.config(text=status_text)
            self.detail_label.config(text=detail_text)
        
        self.root.update_idletasks()
    
    def download_file(self, url, file_path, model_code):
        try:
            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()
            
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            with open(file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if self.is_stopped:
                        return False, "중지됨"
                    if chunk:
                        f.write(chunk)
            
            return True, "성공"
            
        except requests.exceptions.RequestException as e:
            return False, f"네트워크 오류: {str(e)}"
        except Exception as e:
            return False, f"파일 저장 오류: {str(e)}"
    
    def get_file_extension(self, url):
        parsed_url = urlparse(url)
        path = parsed_url.path
        if '.' in path:
            return Path(path).suffix
        return '.bin'  # 기본 확장자
    
    def save_result_excel(self):
        if not self.result_data:
            return
        
        try:
            result_df = pd.DataFrame(self.result_data, columns=['모델코드', 'URL', '파일명', '상태', '메시지', '처리시간'])
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            result_file = f"download_result_{timestamp}.xlsx"
            
            result_df.to_excel(result_file, index=False)
            self.log_message(f"결과 파일 저장됨: {result_file}")
            
        except Exception as e:
            self.log_message(f"결과 저장 실패: {str(e)}")
    
    def start_download(self):
        # 입력 검증
        if not self.file_path_var.get():
            messagebox.showerror("오류", "엑셀 파일을 선택해주세요.")
            return
        
        if not os.path.exists(self.file_path_var.get()):
            messagebox.showerror("오류", "선택한 엑셀 파일이 존재하지 않습니다.")
            return
        
        try:
            self.start_row, self.start_col = self.parse_cell_reference(self.start_cell_var.get())
        except ValueError as e:
            messagebox.showerror("오류", f"올바른 셀 참조를 입력해주세요: {str(e)}")
            return
        
        try:
            delay_ms = int(self.delay_var.get())
            if delay_ms < 0:
                raise ValueError("딜레이는 0 이상이어야 합니다")
        except ValueError:
            messagebox.showerror("오류", "올바른 딜레이 값(밀리초)을 입력해주세요.")
            return
        
        # 다운로드 폴더 생성 확인
        download_folder = self.download_folder_var.get()
        try:
            os.makedirs(download_folder, exist_ok=True)
        except Exception as e:
            messagebox.showerror("오류", f"다운로드 폴더 생성 실패: {str(e)}")
            return
        
        self.is_stopped = False
        self.is_paused = False
        self.result_data = []
        
        # 버튼 상태 변경
        self.start_btn.config(state="disabled")
        self.pause_btn.config(state="normal")
        self.stop_btn.config(state="normal")
        
        # 다운로드 스레드 시작
        self.current_thread = threading.Thread(target=self.download_process)
        self.current_thread.daemon = True
        self.current_thread.start()
    
    def pause_download(self):
        self.is_paused = True
        self.pause_btn.config(state="disabled")
        self.resume_btn.config(state="normal")
        self.log_message("다운로드 일시정지됨")
    
    def resume_download(self):
        self.is_paused = False
        self.pause_btn.config(state="normal")
        self.resume_btn.config(state="disabled")
        self.log_message("다운로드 재개됨")
    
    def stop_download(self):
        self.is_stopped = True
        self.is_paused = False
        
        # 버튼 상태 초기화
        self.start_btn.config(state="normal")
        self.pause_btn.config(state="disabled")
        self.resume_btn.config(state="disabled")
        self.stop_btn.config(state="disabled")
        
        self.log_message("다운로드 중지됨")
        self.save_result_excel()
    
    def download_process(self):
        # 하위 클래스에서 구현
        pass
    
    def run(self):
        self.root.mainloop()
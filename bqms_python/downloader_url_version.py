import pandas as pd
import requests
import os
import time
from datetime import datetime
from file_downloader_base import FileDownloaderBase
from tkinter import ttk, messagebox
import tkinter as tk


class URLFileDownloader(FileDownloaderBase):
    def __init__(self):
        super().__init__()
        self.root.title("URL 버전 파일 다운로더")
        
        # URL 컬럼 선택 UI 추가
        self.setup_url_ui()
    
    def setup_url_ui(self):
        # 메인 프레임을 찾아서 URL 프레임 추가
        main_frame = None
        for child in self.root.winfo_children():
            if isinstance(child, tk.Canvas):
                for grandchild in child.winfo_children():
                    if isinstance(grandchild, ttk.Frame):
                        for ggchild in grandchild.winfo_children():
                            if isinstance(ggchild, ttk.Frame) and str(ggchild).endswith('!frame'):
                                main_frame = ggchild
                                break
        
        if not main_frame:
            return
        
        # URL 컬럼 설정 프레임
        url_frame = ttk.LabelFrame(main_frame, text="📊 컬럼 설정", 
                                  style='Card.TLabelframe', padding="15")
        url_frame.pack(fill="x", pady=(0, 15))
        
        col_row = ttk.Frame(url_frame)
        col_row.pack(fill="x", pady=5)
        
        ttk.Label(col_row, text="모델코드 컬럼:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 5))
        self.model_code_col_var = tk.StringVar(value="A")
        ttk.Entry(col_row, textvariable=self.model_code_col_var, width=5, font=('Arial', 9)).pack(side="left", padx=(0, 20))
        
        ttk.Label(col_row, text="URL 컬럼:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 5))
        self.url_col_var = tk.StringVar(value="B")
        ttk.Entry(col_row, textvariable=self.url_col_var, width=5, font=('Arial', 9)).pack(side="left", padx=(0, 20))
        
        ttk.Label(col_row, text="(예: A, B, C... 또는 0, 1, 2...)", foreground="gray").pack(side="left")
    
    def get_column_index(self, col_identifier):
        """컬럼 식별자를 숫자 인덱스로 변환"""
        if col_identifier.isdigit():
            return int(col_identifier)
        else:
            # A=0, B=1, C=2, ...
            return ord(col_identifier.upper()) - ord('A')
    
    def download_process(self):
        try:
            # 엑셀 파일 읽기
            self.log_message("엑셀 파일을 읽는 중...")
            df = pd.read_excel(self.file_path_var.get())
            
            # 컬럼 인덱스 변환
            model_code_col = self.get_column_index(self.model_code_col_var.get())
            url_col = self.get_column_index(self.url_col_var.get())
            
            # 시작 행부터 처리
            df = df.iloc[self.start_row:]
            
            self.total_count = len(df)
            self.current_index = self.start_row
            self.log_message(f"총 {self.total_count}개의 파일 다운로드 시작")
            
            download_folder = self.download_folder_var.get()
            os.makedirs(download_folder, exist_ok=True)
            
            # 각 행 처리
            for idx, row in df.iterrows():
                if self.is_stopped:
                    break
                
                # 일시정지 체크
                while self.is_paused and not self.is_stopped:
                    time.sleep(0.1)
                
                if self.is_stopped:
                    break
                
                try:
                    model_code = str(row.iloc[model_code_col]).strip()
                    url = str(row.iloc[url_col]).strip()
                    
                    if pd.isna(model_code) or pd.isna(url) or not url:
                        self.result_data.append([model_code, url, "", "X", "URL 없음", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                        self.fail_count += 1
                        self.current_index += 1
                        self.update_status()
                        continue
                    
                    # 이미지 없음 URL 체크
                    no_image_url = self.no_image_url_var.get().strip()
                    if no_image_url and url.strip() == no_image_url:
                        self.result_data.append([model_code, url, "", "X", "이미지 없음", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                        self.fail_count += 1
                        self.current_index += 1
                        self.update_status()
                        continue
                    
                    # 파일 확장자 추출
                    file_extension = self.get_file_extension(url)
                    file_name = f"{model_code}{file_extension}"
                    file_path = os.path.join(download_folder, file_name)
                    
                    self.log_message(f"다운로드 중: {model_code}")
                    
                    # 파일 다운로드
                    success, message = self.download_file(url, file_path, model_code)
                    
                    # 딜레이 적용
                    try:
                        delay_ms = int(self.delay_var.get())
                        if delay_ms > 0:
                            time.sleep(delay_ms / 1000.0)
                    except:
                        pass  # 딜레이 설정 오류 시 무시
                    
                    status = "O" if success else "X"
                    self.result_data.append([
                        model_code, 
                        url, 
                        file_name, 
                        status, 
                        message, 
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ])
                    
                    if success:
                        self.success_count += 1
                    else:
                        self.fail_count += 1
                        self.log_message(f"실패: {model_code} - {message}")
                    
                    self.current_index += 1
                    self.update_status()
                    
                    # 매 100개마다 중간 결과 저장
                    if (self.success_count + self.fail_count) % 100 == 0:
                        self.save_intermediate_result()
                
                except Exception as e:
                    self.log_message(f"행 처리 오류: {str(e)}")
                    self.result_data.append([
                        str(row.iloc[model_code_col]) if not pd.isna(row.iloc[model_code_col]) else "", 
                        str(row.iloc[url_col]) if not pd.isna(row.iloc[url_col]) else "", 
                        "", 
                        "X", 
                        f"처리 오류: {str(e)}", 
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    ])
                    self.fail_count += 1
                    self.current_index += 1
                    self.update_status()
            
            if not self.is_stopped:
                self.log_message("모든 다운로드 완료!")
                self.stop_download()
            
        except Exception as e:
            self.log_message(f"전체 프로세스 오류: {str(e)}")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다: {str(e)}")
    
    def save_intermediate_result(self):
        """중간 결과 저장 (100개마다)"""
        try:
            if self.result_data:
                result_df = pd.DataFrame(self.result_data, columns=['모델코드', 'URL', '파일명', '상태', '메시지', '처리시간'])
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                result_file = f"download_result_temp_{timestamp}.xlsx"
                
                result_df.to_excel(result_file, index=False)
                self.log_message(f"중간 결과 저장: {result_file}")
                
        except Exception as e:
            self.log_message(f"중간 결과 저장 실패: {str(e)}")


if __name__ == "__main__":
    app = URLFileDownloader()
    app.run()
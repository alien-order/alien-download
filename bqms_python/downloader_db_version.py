import oracledb
import pandas as pd
import requests
import os
import time
from datetime import datetime
from file_downloader_base import FileDownloaderBase
from tkinter import ttk, messagebox
import tkinter as tk


class DatabaseFileDownloader(FileDownloaderBase):
    def __init__(self):
        super().__init__()
        self.root.title("DB 조회 파일 다운로더")
        
        # DB 연결 정보 추가
        self.setup_db_ui()
        
    def setup_db_ui(self):
        # 메인 프레임을 찾아서 DB 프레임 추가
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
        
        # DB 연결 정보 프레임 추가
        db_frame = ttk.LabelFrame(main_frame, text="🗄️ DB 연결 정보", 
                                 style='Card.TLabelframe', padding="15")
        db_frame.pack(fill="x", pady=(0, 15))
        
        # 첫 번째 행
        row1 = ttk.Frame(db_frame)
        row1.pack(fill="x", pady=5)
        
        ttk.Label(row1, text="호스트:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 5))
        self.db_host_var = tk.StringVar(value="localhost")
        ttk.Entry(row1, textvariable=self.db_host_var, width=15, font=('Arial', 9)).pack(side="left", padx=(0, 20))
        
        ttk.Label(row1, text="포트:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 5))
        self.db_port_var = tk.StringVar(value="1521")
        ttk.Entry(row1, textvariable=self.db_port_var, width=8, font=('Arial', 9)).pack(side="left", padx=(0, 20))
        
        ttk.Label(row1, text="SID:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 5))
        self.db_sid_var = tk.StringVar(value="XE")
        ttk.Entry(row1, textvariable=self.db_sid_var, width=10, font=('Arial', 9)).pack(side="left")
        
        # 두 번째 행
        row2 = ttk.Frame(db_frame)
        row2.pack(fill="x", pady=5)
        
        ttk.Label(row2, text="사용자:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 5))
        self.db_user_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.db_user_var, width=15, font=('Arial', 9)).pack(side="left", padx=(0, 20))
        
        ttk.Label(row2, text="비밀번호:", font=('Arial', 9, 'bold')).pack(side="left", padx=(0, 5))
        self.db_password_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.db_password_var, width=15, show="*", font=('Arial', 9)).pack(side="left", padx=(0, 20))
        
        # 테스트 연결 버튼
        ttk.Button(row2, text="🔗 연결 테스트", command=self.test_db_connection,
                  style='Action.TButton').pack(side="right")
    
    def test_db_connection(self):
        try:
            dsn = f"{self.db_host_var.get()}:{self.db_port_var.get()}/{self.db_sid_var.get()}"
            
            connection = oracledb.connect(
                user=self.db_user_var.get(),
                password=self.db_password_var.get(),
                dsn=dsn
            )
            
            connection.close()
            messagebox.showinfo("성공", "DB 연결이 성공했습니다!")
            self.log_message("DB 연결 테스트 성공")
            
        except oracledb.Error as e:
            error_msg = f"DB 연결 실패: {str(e)}"
            messagebox.showerror("오류", error_msg)
            self.log_message(error_msg)
    
    def get_db_connection(self):
        dsn = f"{self.db_host_var.get()}:{self.db_port_var.get()}/{self.db_sid_var.get()}"
        
        return oracledb.connect(
            user=self.db_user_var.get(),
            password=self.db_password_var.get(),
            dsn=dsn
        )
    
    def query_urls_from_db(self, model_codes):
        """
        모델코드 리스트를 받아서 DB에서 URL을 조회
        실제 테이블명과 컬럼명은 환경에 맞게 수정 필요
        """
        try:
            connection = self.get_db_connection()
            cursor = connection.cursor()
            
            # IN 절을 위한 placeholder 생성
            placeholders = ','.join([':' + str(i) for i in range(len(model_codes))])
            
            # 실제 SQL은 환경에 맞게 수정
            query = f"""
            SELECT MODEL_CODE, FILE_URL 
            FROM MODEL_FILES 
            WHERE MODEL_CODE IN ({placeholders})
            """
            
            # 딕셔너리 형태로 바인딩
            bind_vars = {str(i): model_codes[i] for i in range(len(model_codes))}
            
            cursor.execute(query, bind_vars)
            results = cursor.fetchall()
            
            cursor.close()
            connection.close()
            
            return {row[0]: row[1] for row in results}  # {model_code: url}
            
        except oracledb.Error as e:
            self.log_message(f"DB 조회 오류: {str(e)}")
            return {}
    
    def download_process(self):
        try:
            # 엑셀 파일 읽기
            self.log_message("엑셀 파일을 읽는 중...")
            df = pd.read_excel(self.file_path_var.get())
            
            # 시작 셀부터 처리
            df = df.iloc[self.start_row:]
            model_codes = df.iloc[:, self.start_col].tolist()  # 시작 셀의 컬럼에서 모델코드 추출
            
            self.total_count = len(model_codes)
            self.current_index = self.start_row
            self.log_message(f"총 {self.total_count}개의 모델코드 처리 시작")
            
            download_folder = self.download_folder_var.get()
            os.makedirs(download_folder, exist_ok=True)
            
            # 1000개씩 배치 처리
            batch_size = 1000
            
            for i in range(0, len(model_codes), batch_size):
                if self.is_stopped:
                    break
                
                # 일시정지 체크
                while self.is_paused and not self.is_stopped:
                    time.sleep(0.1)
                
                if self.is_stopped:
                    break
                
                batch_codes = model_codes[i:i+batch_size]
                batch_start_time = time.time()
                
                self.log_message(f"배치 {i//batch_size + 1} 처리 중... ({len(batch_codes)}개)")
                
                # DB에서 URL 조회
                url_mapping = self.query_urls_from_db(batch_codes)
                
                if not url_mapping:
                    self.log_message(f"배치 {i//batch_size + 1}: DB에서 URL을 찾을 수 없음")
                    # 실패로 기록
                    for code in batch_codes:
                        self.result_data.append([code, "", "", "X", "URL 없음", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                        self.fail_count += 1
                        self.current_index += 1
                    continue
                
                # 각 파일 다운로드
                for model_code in batch_codes:
                    if self.is_stopped:
                        break
                    
                    while self.is_paused and not self.is_stopped:
                        time.sleep(0.1)
                    
                    if self.is_stopped:
                        break
                    
                    url = url_mapping.get(model_code)
                    
                    if not url:
                        # URL이 없는 경우
                        self.result_data.append([model_code, "", "", "X", "URL 없음", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
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
                
                batch_time = time.time() - batch_start_time
                self.log_message(f"배치 {i//batch_size + 1} 완료 ({batch_time:.1f}초)")
            
            if not self.is_stopped:
                self.log_message("모든 다운로드 완료!")
                self.stop_download()
            
        except Exception as e:
            self.log_message(f"오류 발생: {str(e)}")
            messagebox.showerror("오류", f"처리 중 오류가 발생했습니다: {str(e)}")


if __name__ == "__main__":
    app = DatabaseFileDownloader()
    app.run()
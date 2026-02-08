# =============================================================================
# main.py - 쿼리 생성기 백엔드 (리팩토링 버전)
# =============================================================================
import eel
import os
import sys
import base64
import tempfile
import pandas as pd
import logging

# --- 로깅 설정 ---
# 파일에 로그를 기록하고, 매 실행 시 덮어씁니다.
logging.basicConfig(
    filename='debug.log',
    level=logging.DEBUG,
    filemode='w',
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# --- 유틸리티 함수 ---
def resource_path(relative_path):
    """ PyInstaller로 패키징되었을 때 리소스 경로를 올바르게 찾기 위한 함수 """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), 'web', relative_path)

# --- Eel 초기화 ---
eel.init(resource_path('.'))

# --- 백엔드 API (Eel로 노출) ---
@eel.expose
def select_excel_file():
    """ '파일 선택' 대화상자를 띄워 사용자가 엑셀 파일을 선택하게 함 """
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()  # tk 창 숨기기
    root.attributes('-topmost', True)  # 항상 위에 표시
    file_path = filedialog.askopenfilename(
        title='엑셀 파일 선택',
        filetypes=[('Excel Files', '*.xlsx *.xls *.xlsm'), ('All Files', '*.*')]
    )
    root.destroy()
    
    if not file_path:
        return {'success': False, 'error': '파일이 선택되지 않았습니다.'}
    
    return get_excel_info(file_path)

@eel.expose
def handle_dropped_file(base64_data, filename):
    """ 드래그앤드롭된 파일을 처리 """
    try:
        file_bytes = base64.b64decode(base64_data)
        # 임시 파일로 저장하여 처리
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, f"query_gen_{filename}")
        with open(temp_path, 'wb') as f:
            f.write(file_bytes)
        logging.info(f"Dropped file '{filename}' saved to temporary path '{temp_path}'")
        return get_excel_info(temp_path)
    except Exception as e:
        logging.error(f"드롭된 파일 처리 중 에러: {e}", exc_info=True)
        return {'success': False, 'error': f"파일 처리 중 에러: {e}"}

def get_excel_info(file_path):
    """ 엑셀 파일의 시트 이름 목록과 첫 번째 시트의 헤더를 가져옴 """
    try:
        xlsx = pd.ExcelFile(file_path)
        sheet_names = xlsx.sheet_names
        if not sheet_names:
            return {'success': False, 'error': '엑셀 파일에 시트가 없습니다.'}
        
        # 첫 번째 시트의 헤더만 읽어옴
        df_headers = pd.read_excel(xlsx, sheet_name=sheet_names[0], nrows=0)
        headers = [str(h) for h in df_headers.columns]
        
        logging.info(f"'{file_path}'에서 시트 정보 읽기 성공: {sheet_names}")
        
        return {
            'success': True,
            'file_path': file_path,
            'sheet_names': sheet_names,
            'headers': headers
        }
    except Exception as e:
        logging.error(f"엑셀 정보 읽기 중 에러: {e}", exc_info=True)
        return {'success': False, 'error': f"엑셀 파일 읽기 실패: {e}"}

@eel.expose
def get_column_stats(file_path, sheet_name, col_index):
    """ 특정 컬럼의 데이터 통계(원본 개수, 고유 개수)를 계산 """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=[col_index], header=0, dtype=str)
        column_name = df.columns[0]
        
        string_list = [s.strip() for s in df[column_name].dropna().astype(str) if s.strip()]
        original_count = len(string_list)
        unique_count = len(dict.fromkeys(string_list))

        return {
            'success': True,
            'original_count': original_count,
            'unique_count': unique_count
        }
    except Exception as e:
        logging.error(f"컬럼 통계 계산 중 에러: {e}", exc_info=True)
        return {'success': False, 'error': f"통계 계산 실패: {e}"}

@eel.expose
def generate_queries(file_path, sheet_name, col_index, query_template, batch_size):
    """ 선택된 컬럼 데이터로 쿼리를 생성 """
    logging.debug(f"--- 쿼리 생성 시작: file='{file_path}', sheet='{sheet_name}', col={col_index} ---")
    try:
        # 필요한 데이터만 읽기
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=[col_index], header=0, dtype=str)
        column_name = df.columns[0]
        
        # 데이터 처리
        string_list = [s.strip() for s in df[column_name].dropna().astype(str) if s.strip()]
        
        # 중복 제거 (순서 유지)
        unique_values = list(dict.fromkeys(string_list))
        
        if not unique_values:
            return {'success': False, 'error': '선택한 컬럼에 유효한 데이터가 없습니다.'}

        # 숫자 형태인지, 문자 형태인지 판단하여 포맷팅
        is_all_numeric = all(s.replace('.', '', 1).isdigit() for s in unique_values)
        
        if is_all_numeric:
            formatted_values = unique_values
        else:
            formatted_values = [f"'{v.replace('\'', '\'\'')}'" for v in unique_values]

        # {VALUES} 플레이스홀더 확인
        if '{VALUES}' not in query_template:
            return {'success': False, 'error': "쿼리 템플릿에 '{VALUES}' 플레이스홀더가 없습니다."}

        prefix, suffix = query_template.split('{VALUES}', 1)

        # 배치 사이즈에 따라 쿼리 분할 생성
        sub_queries = []
        for i in range(0, len(formatted_values), batch_size):
            chunk = formatted_values[i:i + batch_size]
            values_str = ", ".join(chunk)
            sub_queries.append(prefix + values_str + suffix)

        final_query = "\nUNION ALL\n".join(sub_queries)

        logging.info(f"쿼리 생성 완료. 원본 {len(string_list)}개 -> 고유값 {len(unique_values)}개")
        return {
            'success': True, 
            'query': final_query, 
            'original_count': len(string_list), 
            'unique_count': len(unique_values)
        }

    except Exception as e:
        logging.error(f"쿼리 생성 중 예외 발생: {e}", exc_info=True)
        return {'success': False, 'error': f"쿼리 생성 중 에러 발생: {e}"}


if __name__ == '__main__':
    try:
        eel.start(
            'index.html',
            size=(1280, 700),
            port=0,
            mode='chrome',
            cmdline_args=['--disable-gpu']
        )
    except (IOError, OSError):
        logging.info("Chrome not found, falling back to Edge.")
        eel.start(
            'index.html',
            size=(1280, 700),
            port=0,
            mode='edge'
        )
# =============================================================================
# main.py - URL 다운로더 v2 백엔드
# =============================================================================
#
# [이 프로그램이 뭔가요?]
#   엑셀 파일에서 "파일명"과 "URL"을 읽어서 파일을 한번에 다운로드하는
#   데스크톱 프로그램입니다.
#
# [v1과 뭐가 다른가요?]
#   v1: Python(requests)이 직접 파일을 다운로드
#   v2: 브라우저(JS fetch)가 다운로드 → base64로 Python에 전달 → Python이 저장
#       이렇게 하면 브라우저의 세션/쿠키가 유지되어 로그인이 필요한 URL도 다운로드 가능
#
# [어떤 기술을 쓰나요?]
#   - eel        : Python과 웹 화면(HTML)을 연결해주는 라이브러리
#   - pandas     : 엑셀(.xlsx) 파일을 읽는 라이브러리
#   - tkinter    : Windows 파일/폴더 선택 다이얼로그 (Python 기본 내장)
#
# [프로그램 흐름]
#   1. 사용자가 엑셀 파일 선택  (select_excel_file)
#   2. 엑셀 내용 읽기           (read_excel)
#   3. 사용자가 컬럼/폴더 설정  (화면에서 처리)
#   4. 다운로드 시작             (JS가 fetch로 다운로드 → save_file로 Python에 전달)
#   5. 완료 후 결과 엑셀 생성   (finish_download)
# =============================================================================

import eel          # Python ↔ 웹 화면 연결 라이브러리
import os           # 파일/폴더 경로 처리
import sys          # 시스템 관련 (PyInstaller 빌드 시 경로 처리용)
import base64       # base64 디코딩 (JS에서 받은 파일 데이터, 드래그 앤 드롭)
import tempfile     # 임시 파일 저장용
from pathlib import Path                    # 파일 경로/확장자 처리
from urllib.parse import urlparse, unquote  # URL 분석 (확장자 추출용)

import pandas as pd  # 엑셀 파일 읽기


# =============================================================================
# Eel 초기화 - 웹 화면 폴더 등록
# =============================================================================

def resource_path(relative_path):
    """PyInstaller로 빌드된 EXE에서도 파일 경로가 올바르게 잡히도록 처리"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

eel.init(resource_path('web'))


# =============================================================================
# 전역 변수
# =============================================================================

# 원본 엑셀 파일 경로 (결과 엑셀 생성 시 사용)
_source_excel_path = None

# 다운로드 결과 저장용: { row_index: { 'status': '성공'|'실패'|..., 'path': '저장경로' } }
_download_results = {}

# 저장 폴더 경로
_save_folder = None


# =============================================================================
# 파일/폴더 선택 다이얼로그
# =============================================================================

@eel.expose
def select_excel_file():
    """Windows 파일 선택 창을 띄워서 엑셀 파일을 고르게 합니다."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    file_path = filedialog.askopenfilename(
        title='엑셀 파일 선택',
        filetypes=[
            ('Excel Files', '*.xlsx *.xls *.xlsm'),
            ('All Files', '*.*')
        ]
    )

    root.destroy()
    return file_path if file_path else None


@eel.expose
def select_folder():
    """Windows 폴더 선택 창을 띄워서 저장할 폴더를 고르게 합니다."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title='저장 폴더 선택')
    root.destroy()
    return folder_path if folder_path else None


# =============================================================================
# 엑셀 파일 읽기
# =============================================================================

@eel.expose
def read_excel(file_path):
    """엑셀 파일을 열어서 모든 시트의 데이터를 읽어옵니다."""
    try:
        global _source_excel_path
        _source_excel_path = file_path

        xlsx = pd.ExcelFile(file_path)
        sheets = xlsx.sheet_names

        result = {}
        for sheet_name in sheets:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, dtype=str, header=0)
            headers = [str(h) for h in df.columns]
            data = df.fillna('').values.tolist()

            if len(headers) > 0:
                result[sheet_name] = {
                    'headers': headers,
                    'data': data,
                    'row_count': len(data)
                }

        xlsx.close()
        return {'success': True, 'sheets': result, 'sheet_names': sheets}

    except Exception as e:
        return {'success': False, 'error': str(e)}


@eel.expose
def read_excel_from_data(base64_data, filename):
    """드래그 앤 드롭으로 받은 엑셀 파일 데이터를 읽습니다."""
    try:
        file_bytes = base64.b64decode(base64_data)
        temp_path = os.path.join(tempfile.gettempdir(), filename)
        with open(temp_path, 'wb') as f:
            f.write(file_bytes)
        return read_excel(temp_path)
    except Exception as e:
        return {'success': False, 'error': str(e)}


# =============================================================================
# 다운로드 관련 함수 (JS에서 호출)
# =============================================================================

@eel.expose
def init_download(save_folder):
    """
    다운로드 시작 전 초기화. JS에서 다운로드 루프 시작 전에 호출.

    매개변수:
        save_folder: 파일을 저장할 폴더 경로
    """
    global _download_results, _save_folder
    _download_results = {}
    _save_folder = save_folder
    os.makedirs(save_folder, exist_ok=True)
    return {'success': True}


@eel.expose
def save_file(filename, base64_data, url, subfolder='', row_index=0):
    """
    JS에서 fetch()로 다운로드한 파일 데이터를 저장합니다.

    매개변수:
        filename:    저장할 파일명
        base64_data: 파일 내용 (base64 인코딩된 문자열)
        url:         원본 URL (확장자 추출용)
        subfolder:   하위 폴더명 (빈 문자열이면 루트에 저장)
        row_index:   엑셀 행 번호 (결과 엑셀 매칭용)

    반환값:
        { 'success': True/False, 'path': '저장경로', 'error': '에러메시지' }
    """
    global _download_results

    try:
        filename = filename.strip()

        # 확장자 없으면 URL에서 추출
        if not Path(filename).suffix:
            ext = get_extension_from_url(url)
            if ext:
                filename = filename + ext

        # 저장 경로 결정
        if subfolder and subfolder.strip():
            target_folder = os.path.join(_save_folder, sanitize_filename(subfolder.strip()))
            os.makedirs(target_folder, exist_ok=True)
        else:
            target_folder = _save_folder

        save_path = os.path.join(target_folder, sanitize_filename(filename))
        save_path = get_unique_path(save_path)

        # base64 디코딩 → 파일 저장
        file_bytes = base64.b64decode(base64_data)
        with open(save_path, 'wb') as f:
            f.write(file_bytes)

        # 결과 기록
        _download_results[row_index] = {'status': '성공', 'path': save_path}
        return {'success': True, 'path': save_path}

    except Exception as e:
        _download_results[row_index] = {'status': f'실패: {str(e)}', 'path': ''}
        return {'success': False, 'error': str(e)}


@eel.expose
def mark_skipped(row_index, reason='URL없음'):
    """URL이 없거나 건너뛸 항목을 결과에 기록합니다."""
    global _download_results
    _download_results[row_index] = {'status': reason, 'path': ''}
    return {'success': True}


@eel.expose
def mark_failed(row_index, error_message):
    """JS에서 fetch 실패한 항목을 결과에 기록합니다."""
    global _download_results
    _download_results[row_index] = {'status': f'실패: {error_message}', 'path': ''}
    return {'success': True}


@eel.expose
def finish_download():
    """
    다운로드 완료 후 결과 엑셀을 생성합니다.

    반환값:
        { 'success': True, 'result_file': '결과파일명' }
    """
    result_path = _save_result_excel(_save_folder, _download_results)
    result_name = os.path.basename(result_path) if not result_path.startswith('(') else result_path
    return {'success': True, 'result_file': result_name}


# =============================================================================
# 결과 엑셀 생성
# =============================================================================

def _save_result_excel(save_folder, results):
    """원본 엑셀에 '다운로드결과', '저장경로' 컬럼을 추가한 파일을 저장합니다."""
    try:
        df = pd.read_excel(_source_excel_path, dtype=str, header=0)
        df = df.fillna('')

        status_list = []
        path_list = []
        for idx in range(len(df)):
            if idx in results:
                status_list.append(results[idx]['status'])
                path_list.append(results[idx]['path'])
            else:
                status_list.append('')
                path_list.append('')

        df['다운로드결과'] = status_list
        df['저장경로'] = path_list

        original_name = os.path.splitext(os.path.basename(_source_excel_path))[0]
        result_path = os.path.join(save_folder, f'{original_name}_결과.xlsx')
        result_path = get_unique_path(result_path)

        df.to_excel(result_path, index=False)
        return result_path

    except Exception as e:
        print(f'[결과 엑셀 생성 실패] {e}')
        return f'(생성 실패: {e})'


# =============================================================================
# 유틸리티 함수들
# =============================================================================

def get_extension_from_url(url):
    """URL에서 파일 확장자를 추출합니다."""
    try:
        parsed = urlparse(url)
        path = unquote(parsed.path)
        ext = Path(path).suffix.lower()
        if ext and len(ext) <= 5:
            return ext
    except Exception:
        pass
    return None


def sanitize_filename(filename):
    """Windows에서 파일명으로 사용할 수 없는 특수문자를 밑줄(_)로 바꿉니다."""
    invalid_chars = '<>:"/\\|?*'
    for ch in invalid_chars:
        filename = filename.replace(ch, '_')
    return filename.strip()


def get_unique_path(path):
    """같은 이름의 파일이 이미 있으면 번호를 붙여서 겹치지 않게 합니다."""
    if not os.path.exists(path):
        return path

    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(f"{base} ({counter}){ext}"):
        counter += 1
    return f"{base} ({counter}){ext}"


# =============================================================================
# 프로그램 시작점
# =============================================================================

if __name__ == '__main__':
    try:
        eel.start(
            'index.html',
            size=(720, 680),
            port=0,
            mode='default',             # 기본 브라우저 새 탭으로 열기
            cmdline_args=['--disable-gpu']
        )
    except EnvironmentError:
        eel.start(
            'index.html',
            size=(720, 680),
            port=0,
            mode='edge'
        )

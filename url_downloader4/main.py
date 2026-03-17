# =============================================================================
# main.py - URL 다운로더 v3 백엔드
# =============================================================================
#
# [이 프로그램이 뭔가요?]
#   엑셀 파일에서 "파일명"과 "URL"을 읽어서 파일을 한번에 다운로드하는
#   데스크톱 프로그램입니다.
#
# [v1과 뭐가 다른가요?]
#   v1: Python(requests)이 직접 다운로드 (세션 없음)
#   v3: Python(requests)이 다운로드 + 수동 쿠키 입력 지원
#       → 로그인이 필요한 .do URL 등도 다운로드 가능
#
# [어떤 기술을 쓰나요?]
#   - eel             : Python ↔ 웹 화면 연결
#   - pandas          : 엑셀(.xlsx) 읽기
#   - requests        : 파일 다운로드
#   - threading       : 백그라운드 다운로드
#
# [프로그램 흐름]
#   1. 사용자가 엑셀 파일 선택
#   2. 엑셀 내용 읽기
#   3. 컬럼/폴더/쿠키 설정
#   4. 다운로드 시작 (Python requests + 쿠키)
#   5. 결과 엑셀 생성
# =============================================================================

import eel
import os
import sys
import base64
import tempfile
import requests
import threading
import queue
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import urlparse, unquote

import pandas as pd


# =============================================================================
# Eel 초기화
# =============================================================================

def resource_path(relative_path):
    """PyInstaller 빌드 시에도 경로가 올바르게 잡히도록 처리"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

eel.init(resource_path('web'))


# =============================================================================
# 전역 변수
# =============================================================================

download_cancel = False

_pause_event = threading.Event()
_pause_event.set()

_progress_queue = queue.Queue()

_source_excel_path = None

# 쿠키 저장: requests.cookies.RequestsCookieJar
_cookie_jar = requests.cookies.RequestsCookieJar()


def _notify_progress(data):
    """진행 상황을 큐에 넣음"""
    _progress_queue.put(data)


@eel.expose
def get_progress():
    """JS에서 주기적으로 호출. 큐에 쌓인 진행 상황을 전부 꺼내서 반환."""
    updates = []
    while not _progress_queue.empty():
        try:
            updates.append(_progress_queue.get_nowait())
        except queue.Empty:
            break
    return updates


# =============================================================================
# 쿠키 관련 함수
# =============================================================================

@eel.expose
def set_manual_cookies(cookie_string):
    """수동 쿠키 문자열을 파싱하여 저장. 예: "JSESSIONID=abc123; token=xyz" """
    global _cookie_jar
    _cookie_jar = requests.cookies.RequestsCookieJar()

    try:
        count = 0
        # "key=value; key2=value2" 형태를 파싱
        for pair in cookie_string.split(';'):
            pair = pair.strip()
            if not pair or '=' not in pair:
                continue
            key, value = pair.split('=', 1)
            # domain을 지정하지 않아야 모든 URL에 쿠키가 전송됨
            # (domain을 지정하면 requests가 도메인 매칭을 엄격하게 해서 안 보내는 경우 있음)
            _cookie_jar.set(key.strip(), value.strip())
            count += 1

        return {
            'success': True,
            'count': count,
            'message': f'쿠키 {count}개 적용 완료'
        }

    except Exception as e:
        return {
            'success': False,
            'count': 0,
            'message': f'쿠키 파싱 실패: {str(e)}'
        }


@eel.expose
def clear_cookies():
    """저장된 쿠키를 모두 삭제합니다."""
    global _cookie_jar
    _cookie_jar = requests.cookies.RequestsCookieJar()
    return {'success': True, 'message': '쿠키가 초기화되었습니다.'}


@eel.expose
def get_cookie_count():
    """현재 저장된 쿠키 수를 반환합니다."""
    return len(_cookie_jar)


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
# 다운로드 시작 / 취소 / 일시정지
# =============================================================================

@eel.expose
def start_download(items, save_folder, max_workers=4):
    """다운로드를 시작합니다. 별도 스레드에서 실행."""
    global download_cancel
    download_cancel = False
    _pause_event.set()

    thread = threading.Thread(target=_download_worker, args=(items, save_folder, max_workers))
    thread.daemon = True
    thread.start()


@eel.expose
def cancel_download():
    """다운로드 취소"""
    global download_cancel
    download_cancel = True
    _pause_event.set()


@eel.expose
def pause_download():
    """일시정지"""
    _pause_event.clear()


@eel.expose
def resume_download():
    """일시정지 해제"""
    _pause_event.set()


# =============================================================================
# 실제 다운로드 처리 (쿠키 포함)
# =============================================================================

def _download_one(item, save_folder, total, results, lock, counter):
    """파일 1개를 다운로드 (쿠키 포함)"""
    global download_cancel

    row_idx = item.get('rowIndex', 0)
    filename = item['filename'].strip()
    url = item['url'].strip()

    _pause_event.wait()

    if download_cancel:
        return

    # URL 없으면 건너뛰기
    if not url or url.lower() == 'none':
        with lock:
            results[row_idx] = {'status': 'URL없음', 'path': ''}
            counter[0] += 1
            current = counter[0]
        _notify_progress({'current': current, 'total': total, 'status': 'skip',
                          'filename': filename, 'message': f'URL 없음 - 건너뜀: {filename}'})
        return

    # 엑셀 파일명에 확장자가 있는지 확인
    has_extension = bool(Path(filename).suffix)

    # 저장 경로
    subfolder = item.get('folder', '').strip()
    if subfolder:
        target_folder = os.path.join(save_folder, sanitize_filename(subfolder))
        os.makedirs(target_folder, exist_ok=True)
    else:
        target_folder = save_folder

    with lock:
        save_path = os.path.join(target_folder, sanitize_filename(filename))
        save_path = get_unique_path(save_path)

    try:
        response = requests.get(url, timeout=30, stream=True,
                                cookies=_cookie_jar,
                                headers={
                                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                                })
        response.raise_for_status()

        # 엑셀 파일명에 확장자가 없을 때만 서버 응답에서 확장자 추출
        if not has_extension:
            ext = None
            # 1순위: Content-Disposition 헤더
            content_disp = response.headers.get('Content-Disposition', '')
            if content_disp:
                cd_filename = _parse_content_disposition(content_disp)
                if cd_filename:
                    ext = Path(cd_filename).suffix
            # 2순위: URL 경로 (.do 등 웹 확장자 제외)
            if not ext:
                ext = get_extension_from_url(url)
            if ext:
                with lock:
                    save_path = os.path.join(target_folder, sanitize_filename(filename + ext))
                    save_path = get_unique_path(save_path)

        with open(save_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if download_cancel:
                    break
                f.write(chunk)

        if download_cancel:
            try:
                os.remove(save_path)
            except Exception:
                pass
            with lock:
                results[row_idx] = {'status': '취소됨', 'path': ''}
                counter[0] += 1
            return

        # 성공
        with lock:
            results[row_idx] = {'status': '성공', 'path': save_path}
            counter[0] += 1
            current = counter[0]
        _notify_progress({'current': current, 'total': total, 'status': 'done',
                          'filename': filename, 'message': f'완료: {filename}'})

    except Exception as e:
        with lock:
            results[row_idx] = {'status': f'실패: {str(e)}', 'path': ''}
            counter[0] += 1
            current = counter[0]
        _notify_progress({'current': current, 'total': total, 'status': 'error',
                          'filename': filename, 'message': f'에러: {filename} - {str(e)}'})


def _download_worker(items, save_folder, max_workers=4):
    """다운로드 수행 후 결과 엑셀 생성"""
    global download_cancel

    os.makedirs(save_folder, exist_ok=True)
    total = len(items)

    results = {}
    lock = threading.Lock()
    counter = [0]

    cookie_count = len(_cookie_jar)
    _notify_progress({'current': 0, 'total': total, 'status': 'downloading',
                      'filename': '', 'message': f'{total}개 파일 다운로드 시작 (동시 {max_workers}개, 쿠키 {cookie_count}개)'})

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for item in items:
            if download_cancel:
                break
            future = executor.submit(_download_one, item, save_folder, total, results, lock, counter)
            futures.append(future)

        for future in futures:
            future.result()

    if download_cancel:
        _notify_progress({'current': counter[0], 'total': total, 'status': 'cancelled',
                          'filename': '', 'message': '다운로드 취소됨'})

    # 결과 엑셀 생성
    result_path = _save_result_excel(save_folder, results)

    done_count = sum(1 for r in results.values() if r['status'] == '성공')
    _notify_progress({
        'current': total, 'total': total, 'status': 'complete', 'filename': '',
        'message': f'완료 {done_count}/{total}개 | 결과 파일: {os.path.basename(result_path)}'
    })


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
    """URL에서 파일 확장자를 추출합니다. 웹 프레임워크 확장자는 무시."""
    # .do, .action 등은 파일 확장자가 아니라 웹 프레임워크 URL
    WEB_EXTENSIONS = {'.do', '.action', '.jsp', '.asp', '.aspx', '.php', '.html', '.htm'}
    try:
        parsed = urlparse(url)
        path = unquote(parsed.path)
        ext = Path(path).suffix.lower()
        if ext and len(ext) <= 5 and ext not in WEB_EXTENSIONS:
            return ext
    except Exception:
        pass
    return None


def _parse_content_disposition(header_value):
    """
    Content-Disposition 헤더에서 파일명을 추출합니다.
    .do 같은 URL은 확장자가 없으므로 서버가 보내주는 파일명을 사용합니다.

    예: 'attachment; filename="report.pdf"' → 'report.pdf'
        'attachment; filename*=UTF-8''%ED%8C%8C%EC%9D%BC.pdf' → '파일.pdf'
    """
    if not header_value:
        return None

    import re

    # filename*=UTF-8''인코딩된파일명 (RFC 5987)
    match = re.search(r"filename\*=(?:UTF-8|utf-8)''(.+?)(?:;|$)", header_value)
    if match:
        return unquote(match.group(1).strip())

    # filename="파일명"
    match = re.search(r'filename="(.+?)"', header_value)
    if match:
        return match.group(1).strip()

    # filename=파일명 (따옴표 없이)
    match = re.search(r'filename=([^\s;]+)', header_value)
    if match:
        return match.group(1).strip()

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
            size=(720, 780),
            port=0,
            mode='default',
            cmdline_args=['--disable-gpu']
        )
    except EnvironmentError:
        eel.start(
            'index.html',
            size=(720, 780),
            port=0,
            mode='edge'
        )

# =============================================================================
# main.py - URL 다운로더 백엔드
# =============================================================================
#
# [이 프로그램이 뭔가요?]
#   엑셀 파일에서 "파일명"과 "URL"을 읽어서 파일을 한번에 다운로드하는
#   데스크톱 프로그램입니다.
#
# [어떤 기술을 쓰나요?]
#   - eel        : Python과 웹 화면(HTML)을 연결해주는 라이브러리.
#                   이 프로그램의 화면은 web/index.html 이고,
#                   Python 함수를 화면에서 호출할 수 있게 해줍니다.
#                   (예: 화면에서 버튼 클릭 → Python 함수 실행)
#   - pandas     : 엑셀(.xlsx) 파일을 읽는 라이브러리 (DRM 파일도 지원)
#   - requests   : 인터넷에서 파일을 다운로드하는 라이브러리
#   - tkinter    : Windows 파일/폴더 선택 다이얼로그를 띄우는 라이브러리 (Python 기본 내장)
#   - threading  : 다운로드를 백그라운드에서 실행 (화면이 안 멈추게)
#
# [수정하고 싶을 때]
#   - 화면(UI) 수정     → web/index.html 파일을 수정하세요
#   - 다운로드 동작 수정 → 이 파일(main.py)의 _download_worker() 함수를 수정하세요
#   - EXE로 빌드        → build.bat 실행
#
# [프로그램 흐름]
#   1. 사용자가 엑셀 파일 선택  (select_excel_file)
#   2. 엑셀 내용 읽기           (read_excel)
#   3. 사용자가 컬럼/폴더 설정  (화면에서 처리)
#   4. 다운로드 시작             (start_download → _download_worker)
#   5. 진행 상황을 큐에 쌓음      (_notify_progress → JS가 get_progress로 가져감)
# =============================================================================

import eel          # Python ↔ 웹 화면 연결 라이브러리
import os           # 파일/폴더 경로 처리
import sys          # 시스템 관련 (PyInstaller 빌드 시 경로 처리용)
import base64       # 드래그 앤 드롭 시 파일 데이터 디코딩용
import tempfile     # 임시 파일 저장용
import requests     # 인터넷에서 파일 다운로드
import threading    # 백그라운드 스레드 (다운로드 중 화면 안 멈추게)
import queue        # 스레드 안전한 큐 (다운로드 진행 상황 전달용)
from concurrent.futures import ThreadPoolExecutor  # 병렬 다운로드 (여러 파일 동시에)
from pathlib import Path                    # 파일 경로/확장자 처리
from urllib.parse import urlparse, unquote  # URL 분석 (확장자 추출용)

import pandas as pd  # 엑셀 파일 읽기 (openpyxl보다 DRM 걸린 파일도 잘 읽음)


# =============================================================================
# Eel 초기화 - 웹 화면 폴더 등록
# =============================================================================
# Eel은 'web' 폴더 안의 HTML 파일을 화면으로 사용합니다.
# 이 함수는 PyInstaller로 EXE 빌드했을 때도 파일을 찾을 수 있게 경로를 처리합니다.

def resource_path(relative_path):
    """PyInstaller로 빌드된 EXE에서도 파일 경로가 올바르게 잡히도록 처리"""
    # sys._MEIPASS : PyInstaller가 EXE 실행 시 임시로 풀어놓는 폴더 경로
    # 일반 실행(python main.py)에서는 현재 폴더를 기준으로 경로를 잡음
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

# 'web' 폴더를 Eel에 등록 → 이 폴더 안의 index.html이 프로그램 화면이 됩니다
eel.init(resource_path('web'))


# =============================================================================
# 전역 변수
# =============================================================================

# 다운로드 취소 여부를 저장하는 변수
# True로 바꾸면 다운로드 중간에 멈춤
download_cancel = False

# 일시정지 제어용 Event
# .set() = 진행(기본), .clear() = 일시정지, .wait() = 풀릴 때까지 대기
_pause_event = threading.Event()
_pause_event.set()  # 초기 상태: 진행 중 (막히지 않음)

# 다운로드 진행 상황을 전달하는 큐 (스레드 안전)
_progress_queue = queue.Queue()

# 원본 엑셀 파일 경로 (결과 엑셀 생성 시 사용)
_source_excel_path = None


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
# 파일/폴더 선택 다이얼로그
# =============================================================================
# @eel.expose 란?
#   이 데코레이터를 붙이면 웹 화면(index.html)의 JavaScript에서
#   이 Python 함수를 호출할 수 있게 됩니다.
#   예: JavaScript에서 eel.select_excel_file()() 으로 호출 가능

@eel.expose
def select_excel_file():
    """
    Windows 파일 선택 창을 띄워서 엑셀 파일을 고르게 합니다.

    반환값: 선택한 파일의 전체 경로 (예: "C:/Users/사용자/파일.xlsx")
            취소하면 None 반환
    """
    import tkinter as tk
    from tkinter import filedialog

    # tkinter 창을 만들되 보이지 않게 숨김 (파일 선택 창만 띄우기 위해)
    root = tk.Tk()
    root.withdraw()                    # 메인 창 숨기기
    root.attributes('-topmost', True)  # 최상위에 표시 (다른 창 뒤로 안 가게)

    # 파일 선택 다이얼로그 표시
    file_path = filedialog.askopenfilename(
        title='엑셀 파일 선택',
        filetypes=[
            ('Excel Files', '*.xlsx *.xls *.xlsm'),  # 엑셀 파일만 필터
            ('All Files', '*.*')                       # 또는 모든 파일
        ]
    )

    root.destroy()  # tkinter 정리
    return file_path if file_path else None


@eel.expose
def select_folder():
    """
    Windows 폴더 선택 창을 띄워서 저장할 폴더를 고르게 합니다.

    반환값: 선택한 폴더 경로 (예: "C:/Users/사용자/다운로드")
            취소하면 None 반환
    """
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
    """
    엑셀 파일을 열어서 모든 시트의 데이터를 읽어옵니다.

    매개변수:
        file_path: 엑셀 파일 경로 (예: "C:/파일.xlsx")

    반환값 (성공 시):
        {
            'success': True,
            'sheet_names': ['Sheet1', 'Sheet2', ...],   ← 시트 이름 목록
            'sheets': {
                'Sheet1': {
                    'headers': ['이름', 'URL', ...],    ← 첫 번째 행 (컬럼 제목)
                    'data': [['값1', '값2'], ...],      ← 나머지 데이터 행들
                    'row_count': 100                     ← 데이터 행 수 (헤더 제외)
                },
                ...
            }
        }

    반환값 (실패 시):
        { 'success': False, 'error': '에러 메시지' }
    """
    try:
        global _source_excel_path
        _source_excel_path = file_path  # 결과 엑셀 생성 시 원본 데이터를 다시 읽기 위해 저장

        xlsx = pd.ExcelFile(file_path)
        sheets = xlsx.sheet_names  # 모든 시트 이름 목록

        result = {}
        for sheet_name in sheets:
            # 시트를 DataFrame(표 형태)으로 읽기
            # dtype=str : 모든 셀을 문자열로 읽음 (숫자가 깨지지 않게)
            # header=0  : 첫 번째 행을 컬럼 제목(헤더)으로 사용
            df = pd.read_excel(xlsx, sheet_name=sheet_name, dtype=str, header=0)

            # 헤더(컬럼 제목)와 데이터를 리스트로 변환
            headers = [str(h) for h in df.columns]
            data = df.fillna('').values.tolist()  # NaN(빈 셀)을 '' 로 변환

            if len(headers) > 0:
                result[sheet_name] = {
                    'headers': headers,         # 첫 행 = 컬럼 제목 (헤더)
                    'data': data,               # 나머지 = 실제 데이터
                    'row_count': len(data)       # 데이터 행 수
                }

        xlsx.close()
        return {'success': True, 'sheets': result, 'sheet_names': sheets}

    except Exception as e:
        return {'success': False, 'error': str(e)}


@eel.expose
def read_excel_from_data(base64_data, filename):
    """
    드래그 앤 드롭으로 받은 엑셀 파일 데이터를 읽습니다.

    브라우저에서는 보안상 드래그한 파일의 경로를 알 수 없으므로,
    JavaScript에서 파일 내용을 base64로 인코딩해서 보내줍니다.
    이 함수는 그 데이터를 임시 파일로 저장한 뒤 read_excel()로 읽습니다.

    매개변수:
        base64_data: 파일 내용을 base64로 인코딩한 문자열
        filename: 원본 파일명 (예: "data.xlsx")
    """
    try:
        # base64 디코딩 → 임시 파일로 저장
        file_bytes = base64.b64decode(base64_data)
        temp_path = os.path.join(tempfile.gettempdir(), filename)
        with open(temp_path, 'wb') as f:
            f.write(file_bytes)

        # 임시 파일을 read_excel()로 읽기
        return read_excel(temp_path)
    except Exception as e:
        return {'success': False, 'error': str(e)}


# =============================================================================
# 다운로드 시작 / 취소
# =============================================================================

@eel.expose
def start_download(items, save_folder, max_workers=4):
    """
    다운로드를 시작합니다. 별도 스레드에서 실행되어 화면이 멈추지 않습니다.

    매개변수:
        items: 다운로드할 항목 리스트
               [{ 'filename': '저장할파일명', 'url': '다운로드URL' }, ...]
        save_folder: 파일을 저장할 폴더 경로
        max_workers: 동시 다운로드 수 (1~8, 기본값 4)
    """
    global download_cancel
    download_cancel = False  # 취소 상태 초기화
    _pause_event.set()       # 일시정지 해제 (진행 상태)

    # threading.Thread : 백그라운드에서 _download_worker 함수를 실행
    # daemon=True : 메인 프로그램이 종료되면 이 스레드도 같이 종료됨
    thread = threading.Thread(target=_download_worker, args=(items, save_folder, max_workers))
    thread.daemon = True
    thread.start()


@eel.expose
def cancel_download():
    """다운로드 취소 - download_cancel을 True로 바꾸면 _download_worker가 멈춤"""
    global download_cancel
    download_cancel = True
    _pause_event.set()  # 일시정지 중이면 풀어줘야 취소가 진행됨


@eel.expose
def pause_download():
    """일시정지 - 워커들이 다음 파일 시작 전에 멈춤"""
    _pause_event.clear()


@eel.expose
def resume_download():
    """일시정지 해제 - 멈춰있던 워커들이 다시 진행"""
    _pause_event.set()


# =============================================================================
# 실제 다운로드 처리 (백그라운드 스레드에서 실행됨)
# =============================================================================

def _download_one(item, save_folder, total, results, lock, counter):
    """
    파일 1개를 다운로드하는 함수 (ThreadPoolExecutor의 각 워커가 실행)

    매개변수:
        item: {'rowIndex': 행번호, 'filename': 파일명, 'url': URL, 'folder': 하위폴더}
        save_folder: 저장 폴더
        total: 전체 항목 수
        results: 결과 딕셔너리 (Lock으로 보호)
        lock: threading.Lock (results, counter 접근 보호)
        counter: [완료수] 리스트 (Lock으로 보호, 리스트로 감싸서 참조 전달)
    """
    global download_cancel

    row_idx = item.get('rowIndex', 0)
    filename = item['filename'].strip()
    url = item['url'].strip()

    # 일시정지 대기 (pause 상태면 여기서 멈춤, resume되면 계속 진행)
    _pause_event.wait()

    # 취소 확인
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

    # 확장자 처리
    if not Path(filename).suffix:
        ext = get_extension_from_url(url)
        if ext:
            filename = filename + ext

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
        response = requests.get(url, timeout=30, stream=True, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        response.raise_for_status()

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
        # 실패
        with lock:
            results[row_idx] = {'status': f'실패: {str(e)}', 'path': ''}
            counter[0] += 1
            current = counter[0]
        _notify_progress({'current': current, 'total': total, 'status': 'error',
                          'filename': filename, 'message': f'에러: {filename} - {str(e)}'})


def _download_worker(items, save_folder, max_workers=4):
    """
    다운로드 수행 후 결과 엑셀을 생성하는 함수 (별도 스레드에서 실행됨)

    ThreadPoolExecutor를 사용해 최대 max_workers개 파일을 동시에 다운로드합니다.
    각 항목의 결과를 results 딕셔너리에 기록하고,
    완료 후 원본 엑셀에 '결과', '저장경로' 컬럼을 추가한 결과 파일을 생성합니다.
    """
    global download_cancel

    os.makedirs(save_folder, exist_ok=True)
    total = len(items)

    # 결과 기록용: { rowIndex: { 'status': '성공'|'실패'|..., 'path': '저장경로' } }
    results = {}
    lock = threading.Lock()   # results, counter 접근 보호
    counter = [0]              # 완료 카운터 (리스트로 감싸서 참조 전달)

    _notify_progress({'current': 0, 'total': total, 'status': 'downloading',
                      'filename': '', 'message': f'{total}개 파일 다운로드 시작 (동시 {max_workers}개)'})

    # max_workers개 워커로 병렬 다운로드
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for item in items:
            if download_cancel:
                break
            future = executor.submit(_download_one, item, save_folder, total, results, lock, counter)
            futures.append(future)

        # 모든 작업 완료 대기
        for future in futures:
            future.result()

    if download_cancel:
        _notify_progress({'current': counter[0], 'total': total, 'status': 'cancelled',
                          'filename': '', 'message': '다운로드 취소됨'})

    # ── 결과 엑셀 생성 ──
    result_path = _save_result_excel(save_folder, results)

    done_count = sum(1 for r in results.values() if r['status'] == '성공')
    _notify_progress({
        'current': total, 'total': total, 'status': 'complete', 'filename': '',
        'message': f'완료 {done_count}/{total}개 | 결과 파일: {os.path.basename(result_path)}'
    })


def _save_result_excel(save_folder, results):
    """
    원본 엑셀을 읽어서 오른쪽에 '결과', '저장경로' 컬럼을 추가한 파일을 저장합니다.

    저장 위치: 다운로드 폴더/원본파일명_결과.xlsx
    """
    try:
        # 원본 엑셀의 첫 번째 시트를 읽기
        df = pd.read_excel(_source_excel_path, dtype=str, header=0)
        df = df.fillna('')

        # 결과 컬럼 추가 (각 행의 rowIndex에 해당하는 결과를 매칭)
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

        # 결과 파일 경로: 다운로드폴더/원본파일명_결과.xlsx
        original_name = os.path.splitext(os.path.basename(_source_excel_path))[0]
        result_path = os.path.join(save_folder, f'{original_name}_결과.xlsx')
        result_path = get_unique_path(result_path)

        df.to_excel(result_path, index=False)
        return result_path

    except Exception as e:
        print(f'[결과 엑셀 생성 실패] {e}')
        return f'(생성 실패: {e})'


# =============================================================================
# 유틸리티 함수들 (보조 기능)
# =============================================================================

def get_extension_from_url(url):
    """
    URL에서 파일 확장자를 추출합니다.

    예시:
        get_extension_from_url('https://example.com/photo.png?size=large')
        → '.png'

        get_extension_from_url('https://example.com/api/image')
        → None  (확장자 없음)
    """
    try:
        parsed = urlparse(url)       # URL을 분해 (도메인, 경로, 파라미터 등)
        path = unquote(parsed.path)  # URL 인코딩 해제 (%20 → 공백 등)
        ext = Path(path).suffix.lower()
        if ext and len(ext) <= 5:    # 확장자가 있고 너무 길지 않으면 (.jpeg 정도까지)
            return ext
    except Exception:
        pass
    return None


def sanitize_filename(filename):
    """
    Windows에서 파일명으로 사용할 수 없는 특수문자를 밑줄(_)로 바꿉니다.

    금지 문자: < > : " / \\ | ? *
    예시: 'photo:2024/01.jpg' → 'photo_2024_01.jpg'
    """
    invalid_chars = '<>:"/\\|?*'
    for ch in invalid_chars:
        filename = filename.replace(ch, '_')
    return filename.strip()


def get_unique_path(path):
    """
    같은 이름의 파일이 이미 있으면 번호를 붙여서 겹치지 않게 합니다.

    예시:
        사진.jpg 이 이미 있으면 → 사진 (1).jpg
        사진 (1).jpg 도 있으면  → 사진 (2).jpg
    """
    if not os.path.exists(path):
        return path  # 겹치는 파일 없으면 그대로 사용

    base, ext = os.path.splitext(path)  # 'C:/폴더/사진' 과 '.jpg' 로 분리
    counter = 1
    while os.path.exists(f"{base} ({counter}){ext}"):
        counter += 1
    return f"{base} ({counter}){ext}"


# =============================================================================
# 프로그램 시작점
# =============================================================================
# python main.py 로 실행하면 여기서부터 시작됩니다.
# eel.start()가 웹 화면(index.html)을 Chrome 앱 모드로 엽니다.

if __name__ == '__main__':
    try:
        eel.start(
            'index.html',              # 열 HTML 파일 (web 폴더 안에 있음)
            size=(720, 680),            # 창 크기 (가로 720px, 세로 680px) ← 여기서 창 크기 변경 가능
            port=0,                     # 포트 자동 할당 (0 = 비어있는 포트 자동 선택)
            mode='chrome',              # Chrome 브라우저를 앱 모드로 사용
            cmdline_args=['--disable-gpu']  # 일부 PC에서 GPU 관련 오류 방지
        )
    except EnvironmentError:
        # Chrome이 설치되어 있지 않으면 Edge 브라우저로 대체
        eel.start(
            'index.html',
            size=(720, 680),
            port=0,
            mode='edge'
        )

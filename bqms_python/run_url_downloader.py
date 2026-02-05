"""
URL 미리 추출 버전 파일 다운로더 실행 스크립트

사용법:
1. 모델코드와 URL이 포함된 엑셀 파일 선택
2. 모델코드 컬럼과 URL 컬럼 지정 (A, B, C... 또는 0, 1, 2...)
3. 시작 행 번호 설정
4. 다운로드 폴더 지정
5. 시작 버튼 클릭

특징:
- URL이 이미 엑셀에 있는 경우 사용
- 실시간 진행률 표시
- 일시정지/재개/중지 기능
- 100개마다 중간 결과 저장
- 자동 결과 엑셀 파일 생성
"""

from downloader_url_version import URLFileDownloader

if __name__ == "__main__":
    app = URLFileDownloader()
    app.run()
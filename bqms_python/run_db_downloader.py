"""
DB 조회 버전 파일 다운로더 실행 스크립트

사용법:
1. Oracle DB 연결 정보 입력
2. 모델코드가 포함된 엑셀 파일 선택
3. 시작 행 번호 설정
4. 다운로드 폴더 지정
5. 시작 버튼 클릭

특징:
- 1000개씩 배치로 DB 조회
- 실시간 진행률 표시
- 일시정지/재개/중지 기능
- 자동 결과 엑셀 파일 생성
"""

from downloader_db_version import DatabaseFileDownloader

if __name__ == "__main__":
    app = DatabaseFileDownloader()
    app.run()
@echo off
echo 🔍 시스템 준비 상태 확인...
echo.

echo 📦 Python 버전 확인:
python --version
echo.

echo 📦 필수 패키지 설치:
pip install -r requirements.txt
echo.

echo 🧪 시스템 테스트 실행:
python test_system.py
echo.

echo ✅ 준비 완료! 다음 명령으로 실행:
echo python run_db_downloader.py    (DB 버전)
echo python run_url_downloader.py   (URL 버전)

pause
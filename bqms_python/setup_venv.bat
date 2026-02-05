@echo off
echo 가상환경 설정 및 패키지 설치...

echo 1. 가상환경 생성...
python -m venv venv

echo 2. 가상환경 활성화...
call venv\Scripts\activate.bat

echo 3. 패키지 설치...
pip install -r requirements.txt

echo 4. 설치 확인...
python -c "import pandas, openpyxl, requests; print('모든 패키지 설치 완료')"

echo.
echo 사용법:
echo - 가상환경 활성화: venv\Scripts\activate.bat
echo - 프로그램 실행: python run_query_generator.py
echo - 가상환경 종료: deactivate

pause
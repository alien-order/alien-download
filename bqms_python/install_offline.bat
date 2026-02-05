@echo off
echo 📦 폐쇄망 환경 설치 시작...

echo 1. Python 환경 확인...
python --version
if %errorlevel% neq 0 (
    echo ❌ Python이 설치되지 않았습니다.
    echo Python 3.8 이상을 설치해주세요.
    pause
    exit /b 1
)

echo.
echo 2. 오프라인 패키지 설치...
cd offline_packages
for %%f in (*.whl) do (
    echo Installing %%f...
    pip install %%f --no-index --find-links .
)
cd ..

echo.
echo 3. 설치 확인...
python -c "import pandas, openpyxl, requests, pyperclip; print('✅ 모든 패키지 설치 완료')"

if %errorlevel% equ 0 (
    echo.
    echo ✅ 설치 성공! 이제 프로그램을 실행할 수 있습니다.
    echo.
    echo 실행 명령:
    echo python run_query_generator.py
    echo python run_db_downloader.py
    echo python run_url_downloader.py
) else (
    echo ❌ 설치 중 오류 발생
)

pause

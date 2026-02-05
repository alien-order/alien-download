@echo off
echo 🔧 exe 파일 빌드 중...

echo 📦 의존성 설치...
pip install --upgrade pip setuptools
pip install jaraco.text

echo 🏗️ 쿼리 생성기 exe 빌드...
pyinstaller --onefile --windowed ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module scipy ^
    --exclude-module matplotlib ^
    --exclude-module PIL ^
    --exclude-module cv2 ^
    --exclude-module torch ^
    --exclude-module tensorflow ^
    --hidden-import tkinter ^
    --hidden-import openpyxl ^
    --name="QueryGenerator" query_generator.py

if exist "dist\QueryGenerator.exe" (
    echo ✅ 빌드 완료!
    echo 📁 생성된 파일: dist/QueryGenerator.exe
    echo.
    echo 배포용 파일들:
    echo - dist/QueryGenerator.exe (실행 파일)
    echo - test_model_codes_3500.xlsx (테스트용)
) else (
    echo ❌ 빌드 실패 - 수동 실행 사용:
    echo python run_query_generator.py
)

pause
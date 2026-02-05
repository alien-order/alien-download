import os
import subprocess
import sys


def create_offline_package():
    """폐쇄망용 오프라인 패키지 생성"""
    
    print("폐쇄망용 패키지 생성 중...")
    
    # 1. requirements에서 wheel 파일들 다운로드
    print("\n1. 패키지 다운로드 중...")
    os.makedirs("offline_packages", exist_ok=True)
    
    try:
        subprocess.run([
            sys.executable, "-m", "pip", "download", 
            "-r", "requirements.txt", 
            "-d", "offline_packages",
            "--no-deps"  # 의존성 별도 처리
        ], check=True)
        
        # 주요 의존성들도 따로 다운로드
        core_packages = ["numpy", "python-dateutil", "pytz", "six", "et-xmlfile"]
        for pkg in core_packages:
            try:
                subprocess.run([
                    sys.executable, "-m", "pip", "download", 
                    pkg, 
                    "-d", "offline_packages"
                ], check=False)  # 실패해도 계속 진행
            except:
                pass
                
        print("패키지 다운로드 완료")
        
    except subprocess.CalledProcessError as e:
        print(f"패키지 다운로드 실패: {e}")
        return False
    
    # 2. 설치 스크립트 생성
    print("\n2. 설치 스크립트 생성...")
    
    install_script = """@echo off
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
"""
    
    with open("install_offline.bat", "w", encoding="utf-8") as f:
        f.write(install_script)
    
    # 3. 사용 가이드 생성
    print("3. 사용 가이드 생성...")
    
    guide = """# 폐쇄망 환경 사용 가이드

## 폴더 구조
```
alienorder/
├── offline_packages/          # 오프라인 패키지들 (.whl 파일들)
├── install_offline.bat        # 설치 스크립트
├── run_query_generator.py     # 쿼리 생성기
├── run_db_downloader.py       # DB 다운로더
├── run_url_downloader.py      # URL 다운로더
├── test_model_codes_3500.xlsx # 테스트 파일
└── README_OFFLINE.md          # 이 파일

## 🚀 설치 및 실행

### 1단계: 설치
```bash
install_offline.bat
```

### 2단계: 실행
```bash
# 쿼리 생성기
python run_query_generator.py

# DB 다운로더  
python run_db_downloader.py

# URL 다운로더
python run_url_downloader.py
```

## ⚠️ 주의사항

1. **Python 3.8 이상** 필요 (먼저 설치)
2. 전체 폴더를 그대로 복사해서 사용
3. 인터넷 연결 없이도 동작
4. Windows 환경 전용

## 📋 기능

- **쿼리 생성기**: 엑셀에서 10만개 데이터를 읽어 IN절 쿼리 생성
- **DB 다운로더**: Oracle DB에서 URL 조회 후 파일 다운로드  
- **URL 다운로더**: 엑셀에서 URL 직접 읽어서 파일 다운로드

## 🔧 트러블슈팅

설치 실패 시:
1. Python 버전 확인 (3.8 이상)
2. 관리자 권한으로 실행
3. 바이러스 백신 잠시 해제
"""
    
    with open("README_OFFLINE.md", "w", encoding="utf-8") as f:
        f.write(guide)
    
    print("폐쇄망용 패키지 생성 완료!")
    print("\n생성된 파일들:")
    print("- offline_packages/ (wheel 파일들)")
    print("- install_offline.bat (설치 스크립트)")  
    print("- README_OFFLINE.md (사용 가이드)")
    print("\n폐쇄망에서 사용법:")
    print("1. 전체 폴더를 폐쇄망 PC로 복사")
    print("2. install_offline.bat 실행")
    print("3. python run_query_generator.py 실행")


if __name__ == "__main__":
    create_offline_package()
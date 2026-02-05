# 폐쇄망 환경 사용 가이드

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

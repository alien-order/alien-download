import sys
import platform
import os


def check_compatibility():
    """호환성 체크 및 시스템 정보 출력"""
    
    print("=== 시스템 호환성 체크 ===")
    print(f"현재 Python 버전: {sys.version}")
    print(f"플랫폼: {platform.platform()}")
    print(f"아키텍처: {platform.architecture()}")
    print(f"프로세서: {platform.processor()}")
    
    # Python 버전 체크
    major, minor = sys.version_info[:2]
    print(f"\nPython 버전: {major}.{minor}")
    
    if major < 3 or (major == 3 and minor < 8):
        print("경고: Python 3.8 이상 필요")
        return False
    else:
        print("Python 버전 OK")
    
    # 64비트/32비트 체크
    is_64bit = platform.machine().endswith('64')
    print(f"아키텍처: {'64비트' if is_64bit else '32비트'}")
    
    # Windows 버전 체크
    if platform.system() == 'Windows':
        print("Windows 환경")
        win_version = platform.win32_ver()
        print(f"Windows 버전: {win_version[0]} {win_version[1]}")
    else:
        print("Windows가 아닌 환경")
        return False
    
    return True


def create_system_info():
    """현재 시스템 정보를 파일로 저장"""
    
    info = f"""# 패키징 시스템 정보

**빌드 환경:**
- Python: {sys.version}
- 플랫폼: {platform.platform()}
- 아키텍처: {platform.architecture()[0]}

**호환성 요구사항:**
- Python 3.8 이상
- Windows 64비트 권장
- 메모리: 최소 4GB (대용량 엑셀 처리용)

**주의사항:**
- 동일한 Python 버전 사용 권장
- 32비트/64비트 맞춰서 사용
- 바이러스 백신에서 .exe 파일 허용 필요
"""
    
    with open("SYSTEM_INFO.md", "w", encoding="utf-8") as f:
        f.write(info)
    
    print("시스템 정보가 SYSTEM_INFO.md에 저장됨")


if __name__ == "__main__":
    if check_compatibility():
        print("\n호환성 체크 통과!")
        create_system_info()
        print("\n폐쇄망 설치시 확인사항:")
        print("1. Python 3.8+ 설치 여부")
        print("2. 동일한 아키텍처 (64비트/32비트)")
        print("3. Windows 운영체제")
        print("4. 관리자 권한 (필요시)")
    else:
        print("\n호환성 문제 발견!")
        print("폐쇄망 PC에서 문제가 발생할 수 있습니다.")
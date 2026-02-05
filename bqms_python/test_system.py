"""
시스템 테스트 스크립트 - 실행 전 미리 확인용
"""

import sys
import os

def test_imports():
    """필수 라이브러리 import 테스트"""
    print("📦 라이브러리 import 테스트...")
    
    try:
        import pandas as pd
        print("✅ pandas 정상")
    except ImportError as e:
        print(f"❌ pandas 오류: {e}")
        return False
    
    try:
        import requests
        print("✅ requests 정상")
    except ImportError as e:
        print(f"❌ requests 오류: {e}")
        return False
    
    try:
        import oracledb
        print("✅ oracledb 정상")
    except ImportError as e:
        print(f"❌ oracledb 오류: {e}")
        return False
    
    try:
        import tkinter as tk
        from tkinter import ttk
        print("✅ tkinter 정상")
    except ImportError as e:
        print(f"❌ tkinter 오류: {e}")
        return False
    
    return True

def test_file_structure():
    """파일 구조 확인"""
    print("\n📁 파일 구조 테스트...")
    
    required_files = [
        'file_downloader_base.py',
        'downloader_db_version.py', 
        'downloader_url_version.py',
        'run_db_downloader.py',
        'run_url_downloader.py'
    ]
    
    all_ok = True
    for file in required_files:
        if os.path.exists(file):
            print(f"✅ {file}")
        else:
            print(f"❌ {file} 없음")
            all_ok = False
    
    return all_ok

def test_basic_functionality():
    """기본 기능 테스트"""
    print("\n🧪 기본 기능 테스트...")
    
    try:
        from file_downloader_base import FileDownloaderBase
        
        # 셀 참조 파싱 테스트
        base = FileDownloaderBase()
        
        test_cases = [
            ("A2", (1, 0)),
            ("B1", (0, 1)), 
            ("C10", (9, 2)),
            ("AA1", (0, 26))
        ]
        
        for cell_ref, expected in test_cases:
            try:
                result = base.parse_cell_reference(cell_ref)
                if result == expected:
                    print(f"✅ 셀 참조 {cell_ref} -> {result}")
                else:
                    print(f"❌ 셀 참조 {cell_ref} 오류: {result} != {expected}")
                    return False
            except Exception as e:
                print(f"❌ 셀 참조 {cell_ref} 파싱 오류: {e}")
                return False
        
        base.root.destroy()  # GUI 창 닫기
        return True
        
    except Exception as e:
        print(f"❌ 기본 기능 테스트 오류: {e}")
        return False

def main():
    print("🔍 시스템 테스트 시작...\n")
    
    tests = [
        ("라이브러리 import", test_imports),
        ("파일 구조", test_file_structure),
        ("기본 기능", test_basic_functionality)
    ]
    
    all_passed = True
    
    for test_name, test_func in tests:
        if not test_func():
            print(f"\n❌ {test_name} 테스트 실패")
            all_passed = False
        else:
            print(f"\n✅ {test_name} 테스트 통과")
    
    print("\n" + "="*50)
    
    if all_passed:
        print("🎉 모든 테스트 통과! 실행 준비 완료")
        print("\n실행 명령:")
        print("python run_db_downloader.py    # DB 버전")
        print("python run_url_downloader.py   # URL 버전")
    else:
        print("⚠️  일부 테스트 실패 - 문제를 해결한 후 다시 시도")
        print("pip install -r requirements.txt 실행 후 다시 테스트")

if __name__ == "__main__":
    main()
import pandas as pd
import random


def create_test_excel():
    """3500개 모델코드가 포함된 테스트 엑셀 파일 생성"""
    
    # 3500개 모델코드 생성
    model_codes = []
    for i in range(1, 3501):
        model_codes.append(f"MODEL{i:05d}")  # MODEL00001, MODEL00002, ...
    
    # 테스트용 URL 생성 (일부는 이미지 없음 URL)
    urls = []
    for i, code in enumerate(model_codes):
        if i % 50 == 0:  # 50개마다 하나씩 이미지 없음
            urls.append("no_image")
        else:
            # 다양한 확장자로 테스트 URL 생성
            extensions = ['.jpg', '.png', '.gif', '.pdf', '.zip']
            ext = random.choice(extensions)
            urls.append(f"https://example.com/files/{code.lower()}{ext}")
    
    # DataFrame 생성
    df = pd.DataFrame({
        '모델코드': model_codes,
        'URL': urls,
        '설명': [f'{code} 설명' for code in model_codes]
    })
    
    # 엑셀 파일로 저장
    filename = "test_model_codes_3500.xlsx"
    df.to_excel(filename, index=False)
    
    print(f"테스트 엑셀 파일 생성됨: {filename}")
    print(f"데이터 개수: {len(model_codes)}개")
    print(f"컬럼: 모델코드(A), URL(B), 설명(C)")
    print(f"이미지 없음 URL 개수: {len([url for url in urls if url == 'no_image'])}개")
    print(f"시작셀: A2 설정하여 사용")


if __name__ == "__main__":
    create_test_excel()
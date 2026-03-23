import re

def sort_products(product_list):
    # 영문 품명만 추출 및 정렬
    eng = sorted([p for p in product_list if re.match(r"^[A-Za-z]", str(p))])
    # 한글 품명만 추출 및 정렬 (유니코드 한글 초성 기준)
    kor = sorted([p for p in product_list if re.match(r"^[가-힣]", str(p))], key=lambda x: x)
    # 기타(숫자, 특수문자 등)도 필요하면 추가 가능
    return eng + kor
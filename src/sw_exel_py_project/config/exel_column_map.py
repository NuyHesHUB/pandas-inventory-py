# 재고 관리 대장 엑셀 컬럼명 매핑

INVENTORY_COLUMN_MAP = {
    # (A) 날짜
    "date"         : "날 짜",
    # (B) 도착일
    "arrival_date" : "도착일",
    # (C) 품명
    "product_name" : "품  명",
    # (D,E,F,G,H,I) 입고(수입)
    "supplier"     : "공 급 처",
    "in_quantity"  : "수량(Kg)",
    "unit_price"   : "단 가",
    "exchange_rate": "결제환율",
    "krw_amount"   : "원화 금액",
    "amount"       : "금  액",
    # (J,K) 납품(매출)
    "customer"     : "매 출 처",
    "out_quantity" : "수량(Kg)",
    # (L) 재고(kg)
    "current_stock" : "現재고(Kg)",
    # (M) 운송
    "transport"     : "운송",
    # (N) 창고
    "warehouse"     : "창고",
    # (O) 비고
    "note"          : "비  고",
    # (P) 담당
    "manager"       : "담당"
}
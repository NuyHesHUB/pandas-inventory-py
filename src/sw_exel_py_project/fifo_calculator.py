import pandas as pd
import re
from pathlib import Path
import openpyxl
from .config.exel_column_map import INVENTORY_COLUMN_MAP


def _canon(s: object) -> str:
    """공백/괄호/구분기호 제거 + 소문자 정규화"""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    import unicodedata
    s = unicodedata.normalize("NFKC", str(s))
    return re.sub(r"\s+|\(|\)|\[|\]|{|}|/|\\|-|_", "", s).lower()


def _parse_korean_md(s: str, year: int):
    """'MM월 DD일' → Timestamp 변환"""
    if pd.isna(s): 
        return pd.NaT
    if isinstance(s, pd.Timestamp): 
        return s.normalize()
    txt = str(s).strip()
    m = re.match(r"(\d{1,2})\s*월\s*(\d{1,2})\s*일", txt)
    if m:
        try:
            return pd.Timestamp(year=year, month=int(m[1]), day=int(m[2]))
        except: 
            return pd.NaT
    return pd.to_datetime(txt, errors="coerce")


def safe_extract_value(row, col_name):
    """Series에서 안전하게 값 추출"""
    try:
        val = row[col_name]
        if hasattr(val, 'item'):
            return val.item()
        return val
    except:
        return None


class FIFOCalculator:
    def __init__(self, file_path: Path, year: int):
        self.file_path = file_path
        self.year = year
        self.df = self._load_data()
        
    def _load_data(self):
        """엑셀 파일 로드"""
        wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
        matched_sheets = [s for s in wb.sheetnames if str(self.year) in s]
        if not matched_sheets:
            raise ValueError(f"엑셀에 '{self.year}' 시트가 없습니다.")
        
        df = pd.read_excel(
            self.file_path,
            sheet_name=matched_sheets[0],
            engine="openpyxl",
            header=[2, 3],
            dtype=str
        )
        return df
    
    def _find_columns(self):
        """필요한 컬럼들 찾기"""
        # 날짜 컬럼
        date_col = None
        for col in self.df.columns:
            if isinstance(col, tuple):
                if "날" in str(col[0]) and "짜" in str(col[0]):
                    date_col = col
                    break
        
        # 품명 컬럼  
        product_col = None
        for col in self.df.columns:
            if isinstance(col, tuple):
                if "품" in str(col[0]) and "명" in str(col[0]):
                    product_col = col
                    break
        
        # 입고 수량 컬럼들
        in_qty_cols = []
        for col in self.df.columns:
            if isinstance(col, tuple):
                top_level = _canon(str(col[0]))
                if ("입고" in top_level or "수입" in top_level):
                    for level in col[1:]:
                        if "수량" in str(level) or "kg" in str(level).lower():
                            in_qty_cols.append(col)
                            break
        
        # 출고 수량 컬럼들 (납품/매출)
        out_qty_cols = []
        for col in self.df.columns:
            if isinstance(col, tuple):
                top_level = _canon(str(col[0]))
                if ("납품" in top_level or "매출" in top_level):
                    for level in col[1:]:
                        if "수량" in str(level) or "kg" in str(level).lower():
                            out_qty_cols.append(col)
                            break
        
        # 단가 컬럼
        unit_price_col = None
        for col in self.df.columns:
            if isinstance(col, tuple):
                for level in col:
                    if "단" in str(level) and "가" in str(level):
                        unit_price_col = col
                        break
        
        # 환율 컬럼
        exchange_rate_col = None
        for col in self.df.columns:
            if isinstance(col, tuple):
                for level in col:
                    if "환율" in str(level) or "결제환율" in str(level):
                        exchange_rate_col = col
                        break
        
        return {
            'date': date_col,
            'product': product_col,
            'in_qty': in_qty_cols,
            'out_qty': out_qty_cols,
            'unit_price': unit_price_col,
            'exchange_rate': exchange_rate_col
        }
    
    def _extract_inventory_movements(self, product_name: str):
        """특정 제품의 입고/출고 이력 추출"""
        cols = self._find_columns()
        
        # 디버깅: 찾은 컬럼들 출력
        print(f"[디버그] {product_name} - 찾은 컬럼들:")
        for key, val in cols.items():
            print(f"  {key}: {val}")
        
        # 필수 컬럼 확인
        if not cols['date'] or not cols['product']:
            raise ValueError(f"필수 컬럼을 찾을 수 없습니다: date={cols['date']}, product={cols['product']}")
        
        # 날짜 파싱
        self.df['parsed_date'] = self.df[cols['date']].apply(
            lambda x: _parse_korean_md(x, self.year)
        )
        
        # 해당 제품만 필터링 - 안전한 방식
        product_series = self.df[cols['product']].astype(str).str.strip()
        matching_mask = product_series == product_name
        product_df = self.df[matching_mask].copy()
        print(f"[디버그] {product_name} - 필터링된 행 수: {len(product_df)}")
        
        movements = []
        
        for idx, row in product_df.iterrows():
            # 날짜 안전 추출
            date_val = safe_extract_value(row, 'parsed_date')
            if date_val is None or date_val is pd.NaT:
                continue
            if str(date_val).strip() in ['', 'nan', 'NaT']:
                continue
                
            # 입고 수량 계산
            in_qty = 0
            for col in cols['in_qty']:
                val_raw = safe_extract_value(row, col)
                if val_raw is not None:
                    val = str(val_raw).strip()
                    if val and val not in ['nan', '-', '']:
                        try:
                            qty = float(re.sub(r'[^\d\.\-]', '', val))
                            in_qty += qty
                        except:
                            pass
            
            # 출고 수량 계산  
            out_qty = 0
            for col in cols['out_qty']:
                val_raw = safe_extract_value(row, col)
                if val_raw is not None:
                    val = str(val_raw).strip()
                    if val and val not in ['nan', '-', '']:
                        try:
                            qty = float(re.sub(r'[^\d\.\-]', '', val))
                            out_qty += qty
                        except:
                            pass
            
            # 단가와 환율
            unit_price = None
            exchange_rate = None
            
            if cols['unit_price']:
                val_raw = safe_extract_value(row, cols['unit_price'])
                if val_raw is not None:
                    val = str(val_raw).strip()
                    if val and val not in ['nan', '-', '']:
                        try:
                            unit_price = float(re.sub(r'[^\d\.\-]', '', val))
                        except:
                            pass
            
            if cols['exchange_rate']:
                val_raw = safe_extract_value(row, cols['exchange_rate'])
                if val_raw is not None:
                    val = str(val_raw).strip()
                    if val and val not in ['nan', '-', '']:
                        try:
                            exchange_rate = float(re.sub(r'[^\d\.\-]', '', val))
                        except:
                            pass
            
            # 입고 기록
            if in_qty > 0:
                movements.append({
                    'date': date_val,
                    'type': 'in',
                    'quantity': in_qty,
                    'unit_price': unit_price,
                    'exchange_rate': exchange_rate
                })
            
            # 출고 기록
            if out_qty > 0:
                movements.append({
                    'date': date_val,
                    'type': 'out', 
                    'quantity': out_qty,
                    'unit_price': None,
                    'exchange_rate': None
                })
        
        return sorted(movements, key=lambda x: x['date'])
    
    def calculate_fifo(self, product_name: str, start_date: pd.Timestamp, end_date: pd.Timestamp):
        """FIFO 계산"""
        movements = self._extract_inventory_movements(product_name)
        
        # 입고 로트 관리 (날짜순 정렬된 큐)
        inventory_lots = []
        
        # 출고 내역과 할당 결과
        allocations = []
        
        for movement in movements:
            date = movement['date']
            
            if movement['type'] == 'in':
                # 입고: 새로운 로트를 큐에 추가
                inventory_lots.append({
                    'date': date,
                    'remaining_qty': movement['quantity'],
                    'unit_price': movement['unit_price'],
                    'exchange_rate': movement['exchange_rate']
                })
                
            elif movement['type'] == 'out':
                # 출고: 지정된 기간 내의 출고만 처리
                if start_date <= date <= end_date:
                    remaining_out = movement['quantity']
                    
                    # FIFO로 할당
                    while remaining_out > 0 and inventory_lots:
                        # 가장 오래된 로트부터 처리
                        oldest_lot = inventory_lots[0]
                        
                        # 이 로트에서 얼마나 사용할지 결정
                        allocated = min(remaining_out, oldest_lot['remaining_qty'])
                        
                        if allocated > 0:
                            allocations.append({
                                'sale_date': date,
                                'in_date': oldest_lot['date'],
                                'allocated_qty': allocated,
                                'unit_price': oldest_lot['unit_price'],
                                'exchange_rate': oldest_lot['exchange_rate']
                            })
                            
                            # 로트에서 차감
                            oldest_lot['remaining_qty'] -= allocated
                            remaining_out -= allocated
                            
                            # 로트가 모두 소진되면 제거
                            if oldest_lot['remaining_qty'] <= 0:
                                inventory_lots.pop(0)
                    
                    # 할당하지 못한 수량이 있으면 재고 부족 표시
                    if remaining_out > 0:
                        allocations.append({
                            'sale_date': date,
                            'in_date': None,
                            'allocated_qty': -remaining_out,  # 음수로 부족분 표시
                            'unit_price': None,
                            'exchange_rate': None
                        })
        
        return allocations


def calculate_fifo_for_products(file_path: Path, year: int, products_data: list, 
                               start_date: pd.Timestamp, end_date: pd.Timestamp):
    """
    여러 제품에 대해 FIFO 계산
    products_data: [{'product': '제품명', 'quantity': 판매량}, ...]
    """
    calculator = FIFOCalculator(file_path, year)
    
    all_results = []
    
    for product_info in products_data:
        product_name = product_info['product']
        print(f"[처리중] {product_name} | FIFO 계산")
        
        try:
            allocations = calculator.calculate_fifo(product_name, start_date, end_date)
            
            for allocation in allocations:
                all_results.append({
                    'product': product_name,
                    'sale_date': allocation['sale_date'],
                    'in_date': allocation['in_date'],
                    'allocated_qty': allocation['allocated_qty'],
                    'unit_price': allocation['unit_price'],
                    'exchange_rate': allocation['exchange_rate']
                })
                
        except Exception as e:
            print(f"[오류] {product_name}: {e}")
            import traceback
            traceback.print_exc()
    
    return pd.DataFrame(all_results)

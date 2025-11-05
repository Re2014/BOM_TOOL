# utils.py
import re

# --- 正規表現 ---
ref_pattern = re.compile(r'[A-Z]+[0-9]+')
ref_range_pattern = re.compile(r'^([A-Z]+)(\d+)\s*[-~～ー]\s*([A-Z]*)(\d+)$', re.IGNORECASE)

# --- キーワード ---認識されない列があった場合ここに列名を追加することで認識される場合があります。
HEADER_KEYWORDS = {
    'ref': ['部品番号', 'ref des', 'ロケーション番号', 'ref', '記号', 'designator', 'symbol', 'リファレンス', '回路記号', '位置番号', '部品記号','デバイス番号'],
    'part': ['part number', 'メーカー品番', '型番', '型式', '形式', '型格', '定格', 'part', 'value', '品名', 'description', '図番', '名称', 'パート名','識別符号'],
    'mfg': ['メーカー', 'mfg', 'maker', 'manufacturer', '製造元','製造者']
}

# --- 型番からメーカーを推測する関数 ---
def detect_manufacturer(part_number_string):
    pn_upper = part_number_string.upper()
    if pn_upper.startswith(('GRM', 'GCM', 'BLM')): return 'Murata'
    if pn_upper.startswith('CGA'): return 'TDK'
    if pn_upper.startswith('MCR'): return 'Rohm'
    if pn_upper.startswith('CC'): return 'Yageo'
    pn_lower = part_number_string.lower()
    if 'murata' in pn_lower: return 'Murata'
    if 'tdk' in pn_lower: return 'TDK'
    if 'rohm' in pn_lower: return 'Rohm'
    if 'yageo' in pn_lower: return 'Yageo'
    if 'kyocera' in pn_lower: return 'Kyocera'
    return ""
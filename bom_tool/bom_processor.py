# bom_processor.py
import re

# --- 自作モジュールからインポート ---
from utils import (
    ref_pattern, 
    ref_range_pattern, 
    HEADER_KEYWORDS, 
    detect_manufacturer
)

# --- コアロジック 1: 2Dデータからフラットリストを抽出 ---
def extract_flat_list_from_rows(data_2d, cancellation_refs=set()):
    header_map, header_row_index, best_score = {}, -1, 0
    for i, row in enumerate(data_2d[:20]):
        if not isinstance(row, list): continue
        temp_map, used_cols = {}, set()
        for key in ['ref', 'part', 'mfg']:
            for keyword in HEADER_KEYWORDS[key]:
                found = False
                for j, cell in enumerate(row):
                    if j in used_cols: continue
                    if keyword.replace(" ", "") in str(cell).lower().strip().replace(" ", ""):
                        temp_map[key] = j; used_cols.add(j); found = True; break
                if found: break
        score = len(temp_map)
        if score > best_score:
            best_score, header_map, header_row_index = score, temp_map, i
            if best_score == 3: break
    if best_score < 2: return None, "ヘッダー行（「部品番号」と「型番」など）の特定に失敗しました。"

    flat_list, last_valid = [], {}
    start_index = header_row_index + 1 if header_row_index != -1 else 0
    current_refs_from_last_row = []

    for row in data_2d[start_index:]:
        if not isinstance(row, list) or all(c is None or str(c).strip() == "" for c in row): continue
        def get_cell_value(key):
            idx = header_map.get(key)
            return str(row[idx]).strip() if idx is not None and len(row) > idx and row[idx] is not None else ""

        ref_val_raw = get_cell_value('ref')
        part_val_raw = get_cell_value('part')
        mfg_val_raw = get_cell_value('mfg')
        
        is_part_continuation = part_val_raw in ['上↑', '↑', '"']
        is_mfg_continuation = mfg_val_raw in ['上↑', '↑', '"']
        
        if is_part_continuation: part_val_raw = last_valid.get('part', '')
        elif part_val_raw: last_valid['part'] = part_val_raw
        if is_mfg_continuation: mfg_val_raw = last_valid.get('mfg', '')
        elif mfg_val_raw: last_valid['mfg'] = mfg_val_raw

        ref_val = ref_val_raw.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ')

        if ref_val:
            # この行に新しい Ref がある場合
            all_split_parts = [r for r in re.split(r'[,、\s]+', ref_val) if r]
            
            potential_parts = []
            for part in all_split_parts:
                if ref_range_pattern.match(part) or ref_pattern.match(part):
                    potential_parts.append(part)
            
            expanded_refs = []
            for part in potential_parts:
                range_match = ref_range_pattern.match(part)
                if range_match:
                    prefix, start, opt_prefix, end = range_match.groups()
                    if start and end:
                        try:
                            if not opt_prefix or prefix.upper() == opt_prefix.upper():
                                for i in range(int(start), int(end) + 1): 
                                    expanded_refs.append(f"{prefix}{i}")
                            else:
                                expanded_refs.append(part)
                        except ValueError:
                            expanded_refs.append(part)
                    else:
                         expanded_refs.append(part)
                else:
                    expanded_refs.append(part)
            
            current_refs_from_last_row = [r for r in expanded_refs if r and r not in cancellation_refs]
        
        elif not is_part_continuation and not is_mfg_continuation:
            current_refs_from_last_row = []

        part_val_list = [p.strip() for p in part_val_raw.split('\n') if p.strip()]

        if not any(part_val_list) or not current_refs_from_last_row:
            continue

        for part_line in part_val_list:
            part_val = part_line.split()[0] if part_line else ""
            mfg_val = mfg_val_raw if mfg_val_raw else detect_manufacturer(part_val)
            if part_val:
                for r in current_refs_from_last_row:
                    flat_list.append({"ref": r, "part": part_val, "mfg": mfg_val})
    
    return flat_list, None

# --- コアロジック 2: フラットリストを集計 ---
def group_and_finalize_bom(flat_list):
    grouped_map = {}
    for item in flat_list:
        key = f"{item['part']}||{item['mfg']}"
        if key not in grouped_map:
            grouped_map[key] = {'refs': set(), 'part': item['part'], 'mfg': item['mfg']}
        grouped_map[key]['refs'].add(item['ref'])

    final_results = []
    for group in grouped_map.values():
        sorted_refs = sorted(list(group['refs']), key=lambda x: [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', x)])
        final_results.append({'ref': ', '.join(sorted_refs), 'part': group['part'], 'mfg': group['mfg']})

    return final_results
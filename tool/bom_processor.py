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
def extract_flat_list_from_rows(data_2d, cancellation_refs=set(), remove_parentheses=True):
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

        if remove_parentheses:
            # 括弧削除モード: 従来通り、括弧をスペースに置換
            ref_val = ref_val_raw.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ')
        else:
            # 括弧保持モード: RAWの値をそのまま使用
            ref_val = ref_val_raw

        if ref_val:
            # スペース挿入ロジック
            ref_val_spaced_v2 = re.sub(r'([）)])\s*([（(])', r'\1 \2', ref_val)
            ref_val_spaced_v2 = re.sub(r'([）)])\s*([A-Z]+[0-9]+)', r'\1 \2', ref_val_spaced_v2, flags=re.IGNORECASE)
            ref_val_spaced_v2 = re.sub(r'([A-Z]+[0-9]+)\s*([（(])', r'\1 \2', ref_val_spaced_v2, flags=re.IGNORECASE)

            # この行に新しい Ref がある場合
            all_split_parts = [r for r in re.split(r'[,、\s]+', ref_val_spaced_v2) if r]
            
            expanded_refs = []

            if remove_parentheses:
                # --- 括弧削除モード (レンジ展開あり) ---
                potential_parts = []
                for part in all_split_parts:
                    range_match = ref_range_pattern.match(part)
                    ref_match = ref_pattern.match(part)
                    
                    if (range_match) or (ref_match and ref_match.group(0).upper() == part.upper()):
                        potential_parts.append(part)
            
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
            
            else:
                # --- 括弧保持モード (レンジ展開 "あり" に修正) ---
                for part in all_split_parts:
                    # (Q5-Q8) や (R1) のために括弧を外して検証
                    temp_part_for_validation = part.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ').strip()
                    
                    if not temp_part_for_validation:
                        continue 

                    range_match = ref_range_pattern.match(temp_part_for_validation)
                    ref_match = ref_pattern.match(temp_part_for_validation)
                    
                    # ▼▼▼ 修正: レンジ展開ロジックをここに追加 ▼▼▼
                    if range_match:
                        # (Q5-Q8) や Q5-Q8 は、temp_part_for_validation ('Q5-Q8') に基づいて展開
                        prefix, start, opt_prefix, end = range_match.groups()
                        if start and end:
                            try:
                                if not opt_prefix or prefix.upper() == opt_prefix.upper():
                                    for i in range(int(start), int(end) + 1): 
                                        # 展開されたもの (Q5, Q6...) を追加
                                        expanded_refs.append(f"{prefix}{i}")
                                else:
                                    # R1-C5 のような無効なレンジは、元の (R1-C5) を追加
                                    expanded_refs.append(part) 
                            except ValueError:
                                    expanded_refs.append(part)
                        else:
                             expanded_refs.append(part)
                    
                    elif ref_match and ref_match.group(0).upper() == temp_part_for_validation.upper():
                        # レンジではない (R1) や R1 などの場合
                        # 元の (R1) をそのまま追加
                        expanded_refs.append(part) 
                    # ▲▲▲ 修正ここまで ▲▲▲

            
            # --- 共通の除去ロジック ---
            current_refs_from_last_row = []
            upper_cancellation_refs = {ref.upper() for ref in cancellation_refs}
            
            for r in expanded_refs:
                if not r:
                    continue
                
                # ( ) を外した正規化版を作成し、大文字に統一
                normalized_r = r.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ').strip().upper()
                
                if normalized_r and (normalized_r not in upper_cancellation_refs):
                    current_refs_from_last_row.append(r)

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
    
    def sort_key_func(ref_string):
        normalized_ref = ref_string.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ').strip()
        parts = re.split('([0-9]+)', normalized_ref)
        key_parts = []
        for part in parts:
            if part.isdigit():
                key_parts.append(int(part))
            else:
                key_parts.append(part.lower())
        return key_parts

    for group in grouped_map.values():
        sorted_refs = sorted(list(group['refs']), key=sort_key_func)
        final_results.append({'ref': ', '.join(sorted_refs), 'part': group['part'], 'mfg': group['mfg']})

    return final_results
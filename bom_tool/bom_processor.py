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
    best_header_names = {} 

    for i, row in enumerate(data_2d[:20]):
        if not isinstance(row, list): continue
        temp_map, used_cols = {}, set()
        temp_header_names = {} 

        for key in ['ref', 'part', 'mfg']:
            for keyword in HEADER_KEYWORDS[key]:
                found = False
                for j, cell in enumerate(row):
                    if j in used_cols: continue
                    
                    # セルがオブジェクトになったため .value を参照
                    cell_value_str = str(cell.get("value", "")).lower().strip().replace(" ", "")
                    
                    if keyword.replace(" ", "") in cell_value_str:
                        temp_map[key] = j
                        temp_header_names[key] = str(cell.get("value", "")).strip() 
                        used_cols.add(j); found = True; break
                if found: break
        
        score = len(temp_map) # 'ref', 'part', 'mfg' の3つを見つける
        
        if score > best_score:
            best_score = score
            header_map = temp_map 
            header_row_index = i
            best_header_names = temp_header_names 
            
            if best_score == 3: # 3つ全部見つかったら最高
                break # ループ終了
    
    # コア要件（ref と part）が揃っているかチェック
    if 'ref' not in header_map or 'part' not in header_map:
        
        if 'ref' in best_header_names and 'part' not in best_header_names:
            # Ref は見つかったが Part がない
            found_ref_name = best_header_names.get('ref', 'N/A')
            error_msg = f"ヘッダー行の特定に失敗しました。『部品番号』列は \"{found_ref_name}\" として認識しましたが、『型番』列（例: \"Part Number\", \"メーカー品番\"）が見つかりませんでした。"
            return None, error_msg, []
            
        elif 'part' in best_header_names and 'ref' not in best_header_names:
            # Part は見つかったが Ref がない
            found_part_name = best_header_names.get('part', 'N/A')
            error_msg = f"ヘッダー行の特定に失敗しました。『型番』列は \"{found_part_name}\" として認識しましたが、『部品番号』列（例: \"Ref\", \"Symbol\"）が見つかりませんでした。"
            return None, error_msg, []
            
        else:
            # 両方ない (best_score が 0 または 1 のまま)
            error_msg = "ヘッダー行の特定に失敗しました。先頭20行内に、『部品番号』（例: \"Ref\"）と『型番』（例: \"Part Number\"）の両方に一致するキーワードが見つかりませんでした。"
            return None, error_msg, []

    flat_list, last_valid = [], {}
    start_index = header_row_index + 1
    current_refs_from_last_row = []

    cancellation_warnings_set = set() 
    upper_cancellation_refs = {ref.upper() for ref in cancellation_refs}
    part_strike_warnings_set = set() # 型番取り消し線警告用

    for row in data_2d[start_index:]:
        if not isinstance(row, list) or all(c is None or c.get("value", "").strip() == "" for c in row): continue
        
        def get_cell_value(key):
            idx = header_map.get(key)
            return row[idx] if idx is not None and len(row) > idx and row[idx] is not None else {"value": "", "is_struck": False}

        ref_cell_obj = get_cell_value('ref')
        part_cell_obj = get_cell_value('part')
        mfg_cell_obj = get_cell_value('mfg')
        
        ref_val_raw = ref_cell_obj.get("value", "").strip()
        part_val_raw = part_cell_obj.get("value", "").strip()
        mfg_val_raw = mfg_cell_obj.get("value", "").strip()
        
        is_part_continuation = part_val_raw in ['上↑', '↑', '"']
        is_mfg_continuation = mfg_val_raw in ['上↑', '↑', '"']
        
        if is_part_continuation: part_val_raw = last_valid.get('part', '')
        elif part_val_raw: last_valid['part'] = part_val_raw
        if is_mfg_continuation: mfg_val_raw = last_valid.get('mfg', '')
        elif mfg_val_raw: last_valid['mfg'] = mfg_val_raw

        if remove_parentheses:
            ref_val = ref_val_raw.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ')
        else:
            ref_val = ref_val_raw

        if ref_val:
            ref_val_spaced_v2 = re.sub(r'([）)])\s*([（(])', r'\1 \2', ref_val)
            ref_val_spaced_v2 = re.sub(r'([）)])\s*([A-Z]+[0-9]+)', r'\1 \2', ref_val_spaced_v2, flags=re.IGNORECASE)
            ref_val_spaced_v2 = re.sub(r'([A-Z]+[0-9]+)\s*([（(])', r'\1 \2', ref_val_spaced_v2, flags=re.IGNORECASE)

            all_split_parts = [r for r in re.split(r'[,、\s\.\・/]+', ref_val_spaced_v2) if r]
            
            expanded_refs = []

            if remove_parentheses:
                # --- 括弧削除モード (レンジ展開あり) ---
                last_prefix = ""
                prefix_regex = re.compile(r'^([A-Z]+)', re.IGNORECASE) 
                
                for part in all_split_parts:
                    prefix_match = prefix_regex.match(part)
                    if prefix_match:
                        last_prefix = prefix_match.group(1)
                    
                    current_ref = part
                    if part.isdigit() and last_prefix:
                        current_ref = f"{last_prefix}{part}"
                    
                    range_match = ref_range_pattern.match(current_ref)
                    ref_match = ref_pattern.match(current_ref)
                    
                    if range_match:
                        prefix, start, opt_prefix, end = range_match.groups()
                        if start and end:
                            try:
                                if not opt_prefix:
                                    opt_prefix = prefix
                                
                                if prefix.upper() == opt_prefix.upper():
                                    for i in range(int(start), int(end) + 1): 
                                        expanded_refs.append(f"{prefix}{i}")
                                    last_prefix = prefix
                                else:
                                    expanded_refs.append(current_ref)
                            except ValueError:
                                expanded_refs.append(current_ref)
                        else:
                             expanded_refs.append(current_ref)
                    
                    elif ref_match and ref_match.group(0).upper() == current_ref.upper():
                        expanded_refs.append(current_ref)
            
            else:
                # --- 括弧保持モード (レンジ展開 "あり" に修正) ---
                last_prefix = ""
                prefix_regex = re.compile(r'^([A-Z(（]+)', re.IGNORECASE) 
                
                for part in all_split_parts:
                    temp_part_for_validation = part.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ').strip()
                    if not temp_part_for_validation: continue 

                    prefix_match = prefix_regex.match(temp_part_for_validation)
                    if prefix_match:
                        last_prefix = prefix_match.group(1)
                    
                    current_ref = part
                    if temp_part_for_validation.isdigit() and last_prefix:
                        current_ref = re.sub(temp_part_for_validation, f"{last_prefix}{temp_part_for_validation}", part, 1)
                    
                    temp_part_for_validation = current_ref.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ').strip()
                    if not temp_part_for_validation: continue

                    range_match = ref_range_pattern.match(temp_part_for_validation)
                    ref_match = ref_pattern.match(temp_part_for_validation)
                    
                    if range_match:
                        prefix, start, opt_prefix, end = range_match.groups()
                        if start and end:
                            try:
                                if not opt_prefix:
                                    opt_prefix = prefix

                                if not opt_prefix or prefix.upper() == opt_prefix.upper():
                                    for i in range(int(start), int(end) + 1): 
                                        expanded_refs.append(f"{prefix}{i}")
                                    last_prefix = prefix
                                else:
                                    expanded_refs.append(current_ref) 
                            except ValueError:
                                    expanded_refs.append(current_ref)
                        else:
                             expanded_refs.append(current_ref)
                    
                    elif ref_match and ref_match.group(0).upper() == temp_part_for_validation.upper():
                        expanded_refs.append(current_ref) 

            
            # --- 共通の除去ロジック ---
            current_refs_from_last_row = []
            
            for r in expanded_refs:
                if not r:
                    continue
                
                normalized_r = r.replace('(', ' ').replace(')', ' ').replace('（', ' ').replace('）', ' ').strip().upper()
                
                if normalized_r:
                    if normalized_r in upper_cancellation_refs:
                        cancellation_warnings_set.add(normalized_r)
                    else:
                        current_refs_from_last_row.append(r)

        elif not is_part_continuation and not is_mfg_continuation:
            current_refs_from_last_row = []

        # (型番の取り消し線警告チェック)
        if part_cell_obj.get("is_struck", False) and part_val_raw and current_refs_from_last_row:
            refs_str = ", ".join(current_refs_from_last_row)
            warning_msg = f"警告: 部品番号 {refs_str} の 型番 '{part_val_raw}' に取り消し線があります。"
            part_strike_warnings_set.add(warning_msg)

        part_val_list = [p.strip() for p in part_val_raw.split('\n') if p.strip()]

        if not any(part_val_list) or not current_refs_from_last_row:
            continue

        for part_line in part_val_list:
            part_val = part_line.split()[0] if part_line else ""
            mfg_val = mfg_val_raw if mfg_val_raw else detect_manufacturer(part_val)
            if part_val:
                for r in current_refs_from_last_row:
                    flat_list.append({"ref": r, "part": part_val, "mfg": mfg_val})
    
    # Ref除外警告
    cancellation_warnings = [f"除外: 取り消し線のため {ref} を集計から除外しました。" for ref in sorted(list(cancellation_warnings_set))]
    # 型番取り消し線警告
    part_strike_warnings = sorted(list(part_strike_warnings_set))
    
    # すべての警告を結合して返す
    return flat_list, None, cancellation_warnings + part_strike_warnings

# --- コアロジック 2: フラットリストを集計 ---
def group_and_finalize_bom(flat_list):
    # (この関数 (group_and_finalize_bom) は変更なし)
    ref_to_part_map = {}
    grouped_map = {}
    
    for item in flat_list:
        key = f"{item['part']}||{item['mfg']}"
        ref = item['ref']
        if key not in grouped_map:
            grouped_map[key] = {'refs': set(), 'part': item['part'], 'mfg': item['mfg']}
        grouped_map[key]['refs'].add(ref)
        
        if ref not in ref_to_part_map:
            ref_to_part_map[ref] = set()
        ref_to_part_map[ref].add(key)

    warnings = []
    for ref, part_keys in ref_to_part_map.items():
        if len(part_keys) > 1:
            part_list = [k.split('||')[0] for k in part_keys]
            warning_message = f"重複警告: 部品番号 '{ref}' が複数の異なる型番に割り当てられています: [{', '.join(part_list)}]"
            warnings.append(warning_message)

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

    return final_results, warnings
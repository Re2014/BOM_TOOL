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
            
    if best_score < 2:
        found_keys = header_map.keys()
        error_message = ""
        if 'ref' not in found_keys and 'part' not in found_keys:
            if 'mfg' in found_keys:
                error_message = "ヘッダー行の特定に失敗しました。「メーカー」は見つかりましたが、必須の「部品番号」と「型番」の列が見つかりませんでした。"
            else:
                error_message = "ヘッダー行の特定に失敗しました。「部品番号」「型番」のいずれの列も見つかりませんでした。"
        elif 'ref' in found_keys and 'part' not in found_keys:
            error_message = "ヘッダー行の特定に失敗しました。「部品番号」は見つかりましたが、「型番」の列が見つかりませんでした。"
        elif 'part' in found_keys and 'ref' not in found_keys:
            error_message = "ヘッダー行の特定に失敗しました。「型番」は見つかりましたが、「部品番号」の列が見つかりませんでした。"
        else:
            error_message = "ヘッダー行（「部品番号」と「型番」など）の特定に失敗しました。"
        return None, error_message, []

    flat_list, last_valid = [], {}
    start_index = header_row_index + 1 if header_row_index != -1 else 0
    current_refs_from_last_row = []

    last_prefix = "" 
    cancellation_warnings_set = set()
    upper_cancellation_refs = {ref.upper() for ref in cancellation_refs}

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
            last_prefix = ""

        # ▼▼▼ 変更: 型番(part)が空でも Ref を登録するロジックに変更 ▼▼▼
        
        # 1. 処理すべき Ref がなければ、この行はスキップ
        if not current_refs_from_last_row:
            continue

        # 2. 型番(part)を取得
        part_val_list = [p.strip() for p in part_val_raw.split('\n') if p.strip()]

        if not any(part_val_list):
            # 3. 型番(part)が空の場合
            # メーカー列が空でない可能性を考慮 (例: "R1", "", "Murata")
            mfg_val = mfg_val_raw 
            for r in current_refs_from_last_row:
                flat_list.append({"ref": r, "part": "", "mfg": mfg_val})
        
        else:
            # 4. 型番(part)がある場合 (従来のロジック)
            for part_line in part_val_list:
                part_val = part_line.split()[0] if part_line else ""
                # 型番からメーカーを推測、または指定の mfg_val を使用
                mfg_val = mfg_val_raw if mfg_val_raw else detect_manufacturer(part_val)
                
                if part_val: # part_val が (空白などで) 空になっていないか再確認
                    for r in current_refs_from_last_row:
                        flat_list.append({"ref": r, "part": part_val, "mfg": mfg_val})
        
        # ▲▲▲ 変更ここまで ▲▲▲
    
    cancellation_warnings = [f"除外: 取り消し線のため {ref} を集計から除外しました。" for ref in sorted(list(cancellation_warnings_set))]
    
    return flat_list, None, cancellation_warnings

# --- コアロジック 2: フラットリストを集計 ---
def group_and_finalize_bom(flat_list):
    # (この関数は変更なし)
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

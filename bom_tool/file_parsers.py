# file_parsers.py
import re
import csv
import pdfplumber
import xlrd 

# --- 自作モジュールからインポート ---
from utils import ref_pattern

# --- rich_text=True モードで読み込んだExcelセルを処理する ---
def parse_single_excel_sheet_rich_text(sheet):
    data = []
    cancellation_refs = set()
    
    for row in sheet.iter_rows():
        row_data = []
        for cell in row:
            cell_full_text = ""
            
            if cell.value is None:
                row_data.append({"value": "", "is_struck": False})
                continue

            if isinstance(cell.value, list):
                # ■ It's Rich Text
                text_to_cancel = ""
                cell_has_strike = False # セル内に取り消し線があるか
                for run in cell.value:
                    if isinstance(run, str):
                        cell_full_text += run
                    elif hasattr(run, 'text'):
                        if run.text:
                            cell_full_text += run.text
                            if run.font and run.font.strike:
                                text_to_cancel += " " + run.text
                                cell_has_strike = True # 型番警告用にフラグを立てる
                
                row_data.append({"value": cell_full_text, "is_struck": cell_has_strike})
                
                if text_to_cancel:
                    found_refs = ref_pattern.findall(text_to_cancel)
                    for ref in found_refs:
                        cancellation_refs.add(ref)
            
            else:
                # ■ It's a simple value
                cell_full_text = str(cell.value)
                cell_is_struck = False
                if cell.font and cell.font.strike:
                    cell_is_struck = True

                row_data.append({"value": cell_full_text, "is_struck": cell_is_struck})
                
                if cell_is_struck: # Ref除外ロジックは従来通り
                    found_refs = ref_pattern.findall(cell_full_text)
                    for ref in found_refs:
                        cancellation_refs.add(ref)
        
        data.append(row_data)
        
    return data, cancellation_refs

# --- xlrd (.xls) 用のパーサー ---
def parse_single_excel_sheet_xls(sheet, book):
    data = []
    cancellation_refs = set()
    
    fonts = book.font_list
    
    for r_idx in range(sheet.nrows):
        row_data = []
        
        for c_idx in range(sheet.ncols):
            cell = sheet.cell(r_idx, c_idx)
            
            cell_full_text = str(cell.value)
            
            cell_is_struck = False
            if fonts:
                try:
                    xf = book.xf_list[cell.xf_index]
                    font = fonts[xf.font_index]
                    if font.struck_out:
                        cell_is_struck = True
                except Exception as e:
                    pass 
            
            row_data.append({"value": cell_full_text, "is_struck": cell_is_struck})
            
            if cell_is_struck:
                found_refs = ref_pattern.findall(cell_full_text)
                for ref in found_refs:
                    cancellation_refs.add(ref)
        
        data.append(row_data)
        
    return data, cancellation_refs


# --- CSV / TXT パーサー ---
def parse_csv_or_txt(file_stream, delimiters):
    file_stream.seek(0)
    try: text_data = file_stream.read().decode('utf-8')
    except UnicodeDecodeError:
        file_stream.seek(0)
        text_data = file_stream.read().decode('shift_jis', errors='replace')
    lines = text_data.splitlines()
    data_2d = []
    if len(delimiters) == 1: # CSV
        reader = csv.reader(lines)
        for row in reader:
            cleaned_row = []
            for cell in row:
                cleaned_cell = cell.strip().strip('"').strip(',').strip()
                cleaned_row.append({"value": cleaned_cell, "is_struck": False})
            data_2d.append(cleaned_row)
    else: # TXT
        delimiter_regex = '|'.join(delimiters)
        for line in lines:
            split_row = re.split(delimiter_regex, line)
            cleaned_row = []
            for cell in split_row:
                cleaned_cell = cell.strip().strip('"').strip(',').strip()
                cleaned_row.append({"value": cleaned_cell, "is_struck": False})
            data_2d.append(cleaned_row)
    return data_2d

# --- PDF パーサー ---
def parse_pdf(file_stream):
    data_2d = []
    with pdfplumber.open(file_stream) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table: data_2d.extend(table)
            else:
                text = page.extract_text()
                if text:
                    for line in text.split('\n'): data_2d.append(re.split(r'\s{2,}', line))
    cleaned_data_2d = []
    for row in data_2d:
        if isinstance(row, list):
             cleaned_row = []
             for cell in row:
                cell_val = str(cell).strip().strip('"').strip(',').strip() if cell is not None else ""
                cleaned_row.append({"value": cell_val, "is_struck": False})
             cleaned_data_2d.append(cleaned_row)
    return cleaned_data_2d
# app.py
from flask import Flask, request, jsonify, render_template
import openpyxl
import json
import io
import traceback
import xlrd # (★ 1. xlrd をインポート)

# --- 自作モジュールをインポート ---
from file_parsers import (
    parse_single_excel_sheet_rich_text,
    parse_csv_or_txt,
    parse_pdf,
    parse_single_excel_sheet_xls # (★ 2. 新しいパーサーをインポート)
)
from bom_processor import (
    extract_flat_list_from_rows,
    group_and_finalize_bom
)

# Flaskアプリケーションを作成
app = Flask(__name__)

# --- 1. Webページの表示 ---
@app.route('/')
def index():
    return render_template('index.html')

# --- 2. ファイル処理のエンドポイント ---
@app.route('/process', methods=['POST'])
def process_file_endpoint():
    # ... (デバッグログなどはそのまま) ...
    print("--- [DEBUG] /process エンドポイントが POST メソッドで呼び出されました ---")
    
    if 'file' not in request.files: 
        print("--- [DEBUG] エラー: 'file' が request.files に見つかりません ---")
        return jsonify({"error": "ファイルがありません"}), 400
        
    file = request.files['file']
    
    if file.filename == '': 
        print("--- [DEBUG] エラー: ファイル名が空です ---")
        return jsonify({"error": "ファイルが選択されていません"}), 400
    
    print(f"--- [DEBUG] ファイル '{file.filename}' の処理を開始します ---")
    
    filename = file.filename.lower()
    in_memory_file = io.BytesIO(file.read())
    
    remove_parentheses = request.form.get('remove_parentheses', 'true') == 'true'
    
    all_flat_data = []
    all_cancellation_warnings = set()
    individual_results = {}
    
    try:
        # (★ 3. if文を .xlsx と .xls で分岐させる ★)
        if filename.endswith(('.xlsx', '.xls')):
            selected_sheets_json = request.form.get('sheets', '[]')
            selected_sheets = json.loads(selected_sheets_json)
            
            if not selected_sheets:
                return jsonify({"error": "処理するシートが選択されていません。"}), 400

            # --- .xlsx (openpyxl) の処理 ---
            if filename.endswith('.xlsx'):
                try:
                    workbook = openpyxl.load_workbook(in_memory_file, rich_text=True)
                except Exception as e:
                    print(traceback.format_exc())
                    return jsonify({"error": f"Excel (.xlsx) ファイルの読み込みに失敗しました。 (エラー: {e})"}), 500
                
                for sheet_name in selected_sheets:
                    if sheet_name not in workbook.sheetnames:
                        individual_results[sheet_name] = {"error": "指定されたシートが見つかりません。"}
                        continue
                    
                    sheet = workbook[sheet_name]
                    
                    # 既存の .xlsx パーサーを呼ぶ
                    data_2d, cancellation_refs = parse_single_excel_sheet_rich_text(sheet)
                    
                    # (以降の処理は .xls と共通)
                    flat_list, error, cancellation_warnings_list = extract_flat_list_from_rows(data_2d, cancellation_refs, remove_parentheses)
                    
                    if error:
                        individual_results[sheet_name] = {"error": error}
                    else:
                        final_data, duplicate_warnings = group_and_finalize_bom(flat_list)
                        total_warnings = duplicate_warnings + cancellation_warnings_list
                        individual_results[sheet_name] = {"data": final_data, "warnings": total_warnings}
                        all_flat_data.extend(flat_list)
                        all_cancellation_warnings.update(cancellation_warnings_list)

            # --- .xls (xlrd) の処理 ---
            elif filename.endswith('.xls'):
                in_memory_file.seek(0)
                file_contents = in_memory_file.read()
                
                try:
                    # formatting_info=True で取り消し線情報を取得
                    book = xlrd.open_workbook(file_contents=file_contents, formatting_info=True, on_demand=True)
                except Exception as e:
                    print(traceback.format_exc())
                    return jsonify({"error": f".xlsファイルの読み込みに失敗しました。 (エラー: {e})"}), 500
                
                for sheet_name in selected_sheets:
                    try:
                        sheet = book.sheet_by_name(sheet_name)
                    except xlrd.XLRDError:
                        individual_results[sheet_name] = {"error": "指定されたシートが見つかりません。"}
                        continue
                    
                    # 新しい .xls パーサーを呼ぶ
                    data_2d, cancellation_refs = parse_single_excel_sheet_xls(sheet, book)

                    # (以降の処理は .xlsx と共通)
                    flat_list, error, cancellation_warnings_list = extract_flat_list_from_rows(data_2d, cancellation_refs, remove_parentheses)
                    
                    if error:
                        individual_results[sheet_name] = {"error": error}
                    else:
                        final_data, duplicate_warnings = group_and_finalize_bom(flat_list)
                        total_warnings = duplicate_warnings + cancellation_warnings_list
                        individual_results[sheet_name] = {"data": final_data, "warnings": total_warnings}
                        all_flat_data.extend(flat_list)
                        all_cancellation_warnings.update(cancellation_warnings_list)

        else:
            # Excel以外のファイル（PDF, CSV, TXT）の処理
            data_2d, cancellation_refs = [], set()
            
            if filename.endswith('.csv'):
                data_2d = parse_csv_or_txt(in_memory_file, delimiters=[','])
            elif filename.endswith('.txt'):
                data_2d = parse_csv_or_txt(in_memory_file, delimiters=['\t', r'\s{2,}'])
            elif filename.endswith('.pdf'):
                data_2d = parse_pdf(in_memory_file)
            else:
                return jsonify({"error": "対応していないファイル形式です。"}), 400

            if not data_2d: return jsonify({"error": "ファイルからデータを抽出できませんでした。"}), 500
            
            flat_list, error, cancellation_warnings_list = extract_flat_list_from_rows(data_2d, cancellation_refs, remove_parentheses)
            if error: return jsonify({"error": error}), 500
            
            all_flat_data.extend(flat_list)
            all_cancellation_warnings.update(cancellation_warnings_list)
            individual_results = {}

        # 最終集計
        combined_data, combined_duplicate_warnings = group_and_finalize_bom(all_flat_data)
        combined_total_warnings = combined_duplicate_warnings + sorted(list(all_cancellation_warnings))
        
        if not combined_data:
             print("--- [DEBUG] エラー: 最終集計データが空です ---")
             return jsonify({"error": "有効なデータが見つかりませんでした。列の名称を確認してください"}), 500

        print("--- [DEBUG] 処理成功。JSONを返します ---")
        return jsonify({
            "combined": {"data": combined_data, "warnings": combined_total_warnings},
            "individual": individual_results
        })
        
    except Exception as e:
        print("--- [DEBUG] 処理中に 'except' ブロックでエラーが発生しました ---")
        print(traceback.format_exc())
        return jsonify({"error": f"処理中に予期せぬエラーが発生しました: {e}"}), 500

# --- 5. サーバーの起動 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
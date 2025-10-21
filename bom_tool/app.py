# app.py
from flask import Flask, request, jsonify, render_template
import openpyxl
import json
import io
import traceback

# --- 自作モジュールをインポート ---
from file_parsers import (
    parse_single_excel_sheet_rich_text,
    parse_csv_or_txt,
    parse_pdf
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
    if 'file' not in request.files: return jsonify({"error": "ファイルがありません"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "ファイルが選択されていません"}), 400
    
    filename = file.filename.lower()
    in_memory_file = io.BytesIO(file.read())
    
    all_flat_data = []
    individual_results = {}
    
    try:
        if filename.endswith(('.xlsx', '.xls')):
            selected_sheets_json = request.form.get('sheets', '[]')
            selected_sheets = json.loads(selected_sheets_json)
            
            if not selected_sheets:
                return jsonify({"error": "処理するシートが選択されていません。"}), 400

            try:
                workbook = openpyxl.load_workbook(in_memory_file, rich_text=True)
            except Exception as e:
                print(traceback.format_exc())
                return jsonify({"error": f"Excelファイルの読み込みに失敗しました。サポートされている .xlsx 形式か確認してください。 (エラー: {e})"}), 500
            
            for sheet_name in selected_sheets:
                if sheet_name not in workbook.sheetnames:
                    individual_results[sheet_name] = {"error": "指定されたシートが見つかりません。"}
                    continue
                
                sheet = workbook[sheet_name]
                
                # ▼▼▼ パーサーとプロセッサを呼び出し ▼▼▼
                data_2d, cancellation_refs = parse_single_excel_sheet_rich_text(sheet)
                flat_list, error = extract_flat_list_from_rows(data_2d, cancellation_refs)
                
                if error:
                    individual_results[sheet_name] = {"error": error}
                else:
                    individual_results[sheet_name] = group_and_finalize_bom(flat_list)
                    all_flat_data.extend(flat_list)

        else:
            # Excel以外のファイル（PDF, CSV, TXT）の処理
            data_2d, cancellation_refs = [], set()
            
            # ▼▼▼ パーサーを呼び出し ▼▼▼
            if filename.endswith('.csv'):
                data_2d = parse_csv_or_txt(in_memory_file, delimiters=[','])
            elif filename.endswith('.txt'):
                data_2d = parse_csv_or_txt(in_memory_file, delimiters=['\t', r'\s{2,}'])
            elif filename.endswith('.pdf'):
                data_2d = parse_pdf(in_memory_file)
            else:
                return jsonify({"error": "対応していないファイル形式です。"}), 400

            if not data_2d: return jsonify({"error": "ファイルからデータを抽出できませんでした。"}), 500
            
            # ▼▼▼ プロセッサを呼び出し ▼▼▼
            flat_list, error = extract_flat_list_from_rows(data_2d, cancellation_refs)
            if error: return jsonify({"error": error}), 500
            
            all_flat_data.extend(flat_list)
            individual_results = {}

        # ▼▼▼ 最終集計を呼び出し ▼▼▼
        combined_results = group_and_finalize_bom(all_flat_data)
        
        if not combined_results:
             return jsonify({"error": "有効なデータが見つかりませんでした。"}), 500

        return jsonify({
            "combined": combined_results,
            "individual": individual_results
        })
        
    except Exception as e:
        print(traceback.format_exc())
        return jsonify({"error": f"処理中に予期せぬエラーが発生しました: {e}"}), 500

# --- 5. サーバーの起動 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
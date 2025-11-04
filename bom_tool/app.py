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
                
                data_2d, cancellation_refs = parse_single_excel_sheet_rich_text(sheet)
                flat_list, error = extract_flat_list_from_rows(data_2d, cancellation_refs, remove_parentheses)
                
                if error:
                    individual_results[sheet_name] = {"error": error}
                else:
                    # ▼▼▼ 変更 ▼▼▼
                    # group_and_finalize_bom は (data, warnings) のタプルを返すようになった
                    final_data, warnings = group_and_finalize_bom(flat_list)
                    # データを { "data": ..., "warnings": ... } の辞書型で格納
                    individual_results[sheet_name] = {"data": final_data, "warnings": warnings}
                    # ▲▲▲ 変更ここまで ▲▲▲
                    all_flat_data.extend(flat_list)

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
            
            flat_list, error = extract_flat_list_from_rows(data_2d, cancellation_refs, remove_parentheses)
            if error: return jsonify({"error": error}), 500
            
            all_flat_data.extend(flat_list)
            individual_results = {} # Excel以外は individual を使わない

        # ▼▼▼ 最終集計を呼び出し (変更) ▼▼▼
        # combined_results も (data, warnings) のタプルを受け取る
        combined_data, combined_warnings = group_and_finalize_bom(all_flat_data)
        
        # データの有無チェック (data部分を見る)
        if not combined_data:
             print("--- [DEBUG] エラー: 最終集計データが空です ---")
             return jsonify({"error": "有効なデータが見つかりませんでした。"}), 500

        print("--- [DEBUG] 処理成功。JSONを返します ---")
        return jsonify({
            # combined も { "data": ..., "warnings": ... } の辞書型で返す
            "combined": {"data": combined_data, "warnings": combined_warnings},
            "individual": individual_results
        })
        # ▲▲▲ 変更ここまで ▲▲▲
        
    except Exception as e:
        print("--- [DEBUG] 処理中に 'except' ブロックでエラーが発生しました ---")
        print(traceback.format_exc())
        return jsonify({"error": f"処理中に予期せぬエラーが発生しました: {e}"}), 500

# --- 5. サーバーの起動 ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
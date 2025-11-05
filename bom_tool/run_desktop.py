import webview
import json
import csv
import openpyxl # Excel書き出し用にインポート
from app import app # app.pyからFlaskの 'app' オブジェクトをインポート

# 1. Python側の処理をまとめたAPIクラスを定義
class Api:
    def save_file_dialog(self, data_json, file_type):
        """
        JavaScriptから呼び出される関数。
        ファイル保存ダイアログを開き、指定された形式でファイルを保存する。
        """
        window = webview.active_window()
        
        if file_type == 'excel':
            file_types = ('Excel Workbook (*.xlsx)',)
            default_extension = '.xlsx'
        elif file_type == 'csv':
            file_types = ('CSV File (*.csv)',)
            default_extension = '.csv'
        else:
            return {'status': 'error', 'message': 'Unknown file type'}

        # 2. 「名前を付けて保存」ダイアログを表示
        result = window.create_file_dialog(
            webview.SAVE_DIALOG, 
            file_types=file_types,
            save_filename=f"BOM_converted{default_extension}" # デフォルトのファイル名
        )
        
        if not result:
            return {'status': 'cancelled', 'message': 'Save cancelled'}
        
        # 3. ユーザーが選択した保存パスを取得
        save_path = result[0] if isinstance(result, (list, tuple)) else result
        
        # 4. 拡張子を強制
        if not save_path.lower().endswith(default_extension):
            save_path += default_extension

        try:
            # 5. JavaScriptから渡されたJSONデータをPythonの辞書リストに戻す
            data = json.loads(data_json)
            
            # 6. ファイルタイプに応じてPythonでファイルを生成・保存
            if file_type == 'excel':
                self._save_excel(data, save_path)
            elif file_type == 'csv':
                self._save_csv(data, save_path)
                
            return {'status': 'success', 'path': save_path}
        
        except Exception as e:
            # エラーが発生したらJavaScript側にメッセージを返す
            return {'status': 'error', 'message': str(e)}

    # --- Python側でのファイル書き出し処理 ---
    
    def _save_excel(self, data, save_path):
        """
        openpyxlを使ってExcelファイルを作成・保存する
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'BOM'
        
        # ヘッダー
        ws.append(['部品番号', '部品型番', 'メーカー'])
        
        # データ
        for row in data:
            ws.append([row.get('ref', ''), row.get('part', ''), row.get('mfg', '')])
            
        wb.save(save_path) # 指定されたパスに保存

    def _save_csv(self, data, save_path):
        """
        csvモジュールを使ってCSVファイルを作成・保存する
        """
        # utf-8-sig にするとExcelで開いたときに文字化けしない
        with open(save_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow(['部品番号', '部品型番', 'メーカー'])
            
            # データ
            for row in data:
                writer.writerow([row.get('ref', ''), row.get('part', ''), row.get('mfg', '')])

# -----------------
# メインの処理
# -----------------
if __name__ == '__main__':
    api = Api() # 上で定義したAPIクラスのインスタンスを作成
    
    window = webview.create_window(
        'インテリジェントBOM変換ツール',
        app,
        js_api=api # 7. APIインスタンスを 'js_api' として登録
    )
    
    # ▼▼▼ 必ずこの修正を行ってください ▼▼▼
    # http_server=True を追加します
    webview.start(http_server=True, debug=True)
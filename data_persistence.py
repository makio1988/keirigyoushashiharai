#!/usr/bin/env python3
"""
データ永続化ユーティリティ
Render環境でのスリープ対策として、重要なデータを外部に保存・復元
"""
import json
import os
import base64
import gzip
from datetime import datetime

class DataPersistenceManager:
    """データ永続化管理クラス"""
    
    def __init__(self):
        self.backup_dir = "backups"
        os.makedirs(self.backup_dir, exist_ok=True)
    
    def compress_and_encode(self, data):
        """データを圧縮・エンコード"""
        try:
            # JSONに変換
            json_str = json.dumps(data, ensure_ascii=False)
            # UTF-8エンコード
            json_bytes = json_str.encode('utf-8')
            # gzip圧縮
            compressed = gzip.compress(json_bytes)
            # base64エンコード
            encoded = base64.b64encode(compressed).decode('ascii')
            return encoded
        except Exception as e:
            print(f"データ圧縮エラー: {e}")
            return None
    
    def decode_and_decompress(self, encoded_data):
        """データをデコード・展開"""
        try:
            # base64デコード
            compressed = base64.b64decode(encoded_data.encode('ascii'))
            # gzip展開
            json_bytes = gzip.decompress(compressed)
            # UTF-8デコード
            json_str = json_bytes.decode('utf-8')
            # JSONパース
            data = json.loads(json_str)
            return data
        except Exception as e:
            print(f"データ展開エラー: {e}")
            return None
    
    def backup_to_file(self, data, filename):
        """ファイルにバックアップ"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{filename}_{timestamp}.backup"
            backup_path = os.path.join(self.backup_dir, backup_filename)
            
            with open(backup_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            print(f"バックアップ作成: {backup_path}")
            return backup_path
        except Exception as e:
            print(f"バックアップエラー: {e}")
            return None
    
    def restore_from_file(self, filename_pattern):
        """最新のバックアップファイルから復元"""
        try:
            backup_files = [f for f in os.listdir(self.backup_dir) if filename_pattern in f and f.endswith('.backup')]
            if not backup_files:
                return None
            
            # 最新のバックアップファイルを選択
            latest_backup = sorted(backup_files)[-1]
            backup_path = os.path.join(self.backup_dir, latest_backup)
            
            with open(backup_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            print(f"バックアップから復元: {backup_path}")
            return data
        except Exception as e:
            print(f"復元エラー: {e}")
            return None
    
    def auto_backup_payments(self, payments_data):
        """支払データの自動バックアップ"""
        if payments_data:
            self.backup_to_file(payments_data, "payments")
            # 圧縮版も作成
            compressed = self.compress_and_encode(payments_data)
            if compressed:
                with open(os.path.join(self.backup_dir, "payments_compressed.txt"), 'w') as f:
                    f.write(compressed)
    
    def auto_restore_payments(self):
        """支払データの自動復元"""
        # まず通常のバックアップから試行
        data = self.restore_from_file("payments")
        if data:
            return data
        
        # 圧縮版から復元を試行
        try:
            compressed_file = os.path.join(self.backup_dir, "payments_compressed.txt")
            if os.path.exists(compressed_file):
                with open(compressed_file, 'r') as f:
                    compressed_data = f.read()
                data = self.decode_and_decompress(compressed_data)
                if data:
                    print("圧縮バックアップから復元成功")
                    return data
        except Exception as e:
            print(f"圧縮バックアップ復元エラー: {e}")
        
        return None

# グローバルインスタンス
persistence_manager = DataPersistenceManager()

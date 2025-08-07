#!/usr/bin/env python3
"""
日本語フォントをダウンロードするスクリプト
Google Fonts から Noto Sans JP をダウンロード
"""
import urllib.request
import os

def download_japanese_font():
    """日本語フォントをダウンロード"""
    font_url = "https://fonts.gstatic.com/s/notosansjp/v52/nKKF-GM_FYFRJvXzVXaAPe97P1KHynJFP716qHB--oWTiYjNvVA.woff2"
    font_dir = "fonts"
    font_file = os.path.join(font_dir, "NotoSansJP-Regular.woff2")
    
    # フォントディレクトリを作成
    os.makedirs(font_dir, exist_ok=True)
    
    try:
        print(f"日本語フォントをダウンロード中: {font_url}")
        urllib.request.urlretrieve(font_url, font_file)
        print(f"フォントファイルを保存しました: {font_file}")
        return font_file
    except Exception as e:
        print(f"フォントダウンロードエラー: {e}")
        return None

if __name__ == "__main__":
    download_japanese_font()

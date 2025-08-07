#!/usr/bin/env python3
"""
日本語フォント設定スクリプト
ReportLab用のTTFフォントを準備
"""
import urllib.request
import os
import tempfile

def setup_japanese_font():
    """日本語フォント設定"""
    font_dir = "fonts"
    os.makedirs(font_dir, exist_ok=True)
    
    # Noto Sans JP TTF版のURL（GitHub Releases）
    font_urls = [
        "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJK-Regular.otf",
        "https://fonts.gstatic.com/s/notosansjp/v52/nKKF-GM_FYFRJvXzVXaAPe97P1KHynJFP716qHB--oWTiYjNvVA.woff2"
    ]
    
    for i, font_url in enumerate(font_urls):
        try:
            font_file = os.path.join(font_dir, f"NotoSansJP-Regular-{i}.otf")
            print(f"フォントダウンロード試行 {i+1}: {font_url}")
            urllib.request.urlretrieve(font_url, font_file)
            print(f"フォントファイル保存成功: {font_file}")
            return font_file
        except Exception as e:
            print(f"フォントダウンロード失敗 {i+1}: {e}")
            continue
    
    print("全てのフォントダウンロードが失敗しました")
    return None

if __name__ == "__main__":
    setup_japanese_font()

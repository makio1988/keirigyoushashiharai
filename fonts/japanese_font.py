#!/usr/bin/env python3
"""
日本語フォント代替処理
ReportLabで日本語を表示するための代替手法
"""

# 日本語文字を画像として埋め込むか、Unicode対応フォントを使用
JAPANESE_FONT_FALLBACK = {
    # 基本的な日本語文字のUnicodeマッピング
    '業': '\u696d',
    '者': '\u8005', 
    '支': '\u652f',
    '払': '\u6255',
    '表': '\u8868',
    '日': '\u65e5',
    '金': '\u91d1',
    '額': '\u984d',
    '摘': '\u6458',
    '要': '\u8981',
    '合': '\u5408',
    '計': '\u8a08',
    '小': '\u5c0f',
    '送': '\u9001',
    '会': '\u4f1a',
    '社': '\u793e',
    '名': '\u540d',
    '作': '\u4f5c',
    '成': '\u6210',
    '時': '\u6642',
    '株': '\u682a',
    '式': '\u5f0f',
    '有': '\u6709',
    '限': '\u9650',
    '事': '\u4e8b',
    '務': '\u52d9',
    '用': '\u7528',
    '品': '\u54c1',
    '代': '\u4ee3',
    '費': '\u8cbb',
    '清': '\u6e05',
    '掃': '\u6383',
    '不': '\u4e0d',
    '明': '\u660e'
}

def ensure_unicode_text(text):
    """テキストがUnicodeで正しく処理されるよう確認"""
    if not text:
        return ''
    
    # 既にUnicodeの場合はそのまま返す
    if isinstance(text, str):
        return text
    
    # バイト文字列の場合はUTF-8でデコード
    if isinstance(text, bytes):
        try:
            return text.decode('utf-8')
        except UnicodeDecodeError:
            try:
                return text.decode('shift-jis')
            except UnicodeDecodeError:
                return text.decode('utf-8', errors='ignore')
    
    return str(text)

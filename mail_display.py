import streamlit as st
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
import chardet
import re
import mysql.connector
import google.generativeai as genai
from collections import OrderedDict
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lex_rank import LexRankSummarizer
from datetime import timezone, timedelta, datetime
import pytz

# APIキーを設定
genai.configure(api_key="")
# Geminiモデルの設定
model = genai.GenerativeModel('gemini-pro')

# データベースへの接続を定義
host = ""
user = ""
password = ""
database = ""

st.set_page_config(layout="wide")

def delete_all_customers():
    connection = get_connection()
    cursor = connection.cursor()
    # customersテーブルの全てのデータを削除
    cursor.execute("DELETE FROM customers;")
    # 変更をコミット
    connection.commit()
    # 接続のクローズ
    cursor.close()
    connection.close()

# ヘッダー情報のデコード処理
def decode_mime_words(s):
    decoded_fragments = decode_header(s)
    decoded_string = ''
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            if encoding is None:
                encoding = chardet.detect(fragment)['encoding']
            decoded_string += fragment.decode(encoding or 'utf-8', errors='ignore')
        else:
            decoded_string += fragment
    return decoded_string

def extract_email_details(part, subject, sender_name, sender_address, raw_date):
    charset = part.get_content_charset() or 'utf-8'
    email_body = part.get_payload(decode=True).decode(charset)
    return {
        'subject': subject,
        'sender_name': sender_name,
        'sender_address': sender_address,
        'body': email_body,
        'date': raw_date
    }

# メールサーバーに接続し、メールを取得する関数
def fetch_emails(mail_address, mail_password):
    match_imap = {
        'imap-mail': ['@outlook.com', '@hotmail.com'],
        'imap_mail': ['@yahoo.com', '@icloud.com']
    }
    if mail_address in match_imap['imap-mail']:
        imap_server = "imap-mail." + mail_address.split('@')[1]
    elif mail_address in match_imap['imap_mail']:
        imap_server = "imap.mail." + mail_address.split('@')[1]
    else:
        imap_server = "imap." + mail_address.split('@')[1]
    
    server = imaplib.IMAP4_SSL(imap_server)
    server.login(mail_address, mail_password)

    # 受信トレイを選択
    server.select("inbox")

    try:
        # データに保存されている最新の日時から検索
        date_received = MYSQL_fetch(mail_address, 'date_received')
        date_received_utc = date_received - timedelta(hours=9)

        # UTC日時をIMAPフォーマットの文字列に変換
        date_str = date_received_utc.strftime("%d-%b-%Y")
        status, messages = server.search(None, f'SINCE {date_str}')
    except Exception as e:
        # すべてのメールを検索
        status, messages = server.search(None, 'ALL')

    emails = []

    # メールIDを取得
    for num in messages[0].split():
        status, msg_data = server.fetch(num, "(RFC822)")
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])

                # 件名、送信者、受信日をデコード
                raw_subject = msg["subject"]
                subject = decode_mime_words(raw_subject) if raw_subject else "(件名なし)"

                raw_sender = msg.get("From")
                if raw_sender:
                    sender = decode_mime_words(raw_sender)
                    match = re.match(r"(.+?)\s+<(.+?)>", raw_sender)
                    if match:  # matchがNoneでないかを確認
                        sender_name = decode_mime_words(match.group(1))
                        sender_address = decode_mime_words(match.group(2))
                    else:
                        sender_name = ""
                        sender_address = decode_mime_words(raw_sender)  # アドレスだけを取得する
                else:
                    sender_name = ""
                    sender_address = ""

                raw_date = msg.get("Date")
                date_time_utc = parsedate_to_datetime(raw_date)
                date_time_utc = date_time_utc.astimezone(pytz.utc)
                date_to_compare = date_time_utc.strftime("%Y-%m-%d %H:%M:%S")
                date_to_compare = datetime.strptime(date_to_compare, "%Y-%m-%d %H:%M:%S")

                try:
                    if date_received_utc < date_to_compare:
                        # 本文を取得
                        if msg.is_multipart():
                            for part in msg.walk():
                                if part.get_content_type() == "text/plain":
                                    tmp_details = extract_email_details(part, subject, sender_name, sender_address, raw_date)
                                    emails.append(tmp_details)
                        else:
                            tmp_details = extract_email_details(part, subject, sender_name, sender_address, raw_date)
                            emails.append(tmp_details)
                except Exception as e:
                    # 本文を取得
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                tmp_details = extract_email_details(part, subject, sender_name, sender_address, raw_date)
                                emails.append(tmp_details)
                    else:
                        tmp_details = extract_email_details(part, subject, sender_name, sender_address, raw_date)
                        emails.append(tmp_details)
    
    server.close()
    server.logout()
    return emails

# メール本文をgeminiを使って要約する関数
def AI_summary(body):
    prompt = body + '\n' + '文章を要約してください'
    # Gemini APIを使って応答を生成
    response = model.generate_content(prompt)

    # 応答をテキストとして取得（ここではresponse.textと仮定）
    assistant_response = response.text

    return assistant_response

# メール本文を要約する関数
def NLP_summary(subject):
    #要約対象のテキストを指定
    parser = PlaintextParser.from_string(subject, Tokenizer('japanese'))
    #アルゴリズムのインスタンス生成
    summarizer =  LexRankSummarizer()
    #要約の実行 sentences_count で何行に要約したいかを指定する
    res = summarizer(document=parser.document, sentences_count=5)

    return res

# MySQLへの接続設定
def get_connection():
    try:
        connection = mysql.connector.connect(
            host = host,
            user = user,
            password = password,
            database = database  # データベースを指定
        )
        return connection
    except Exception as e:
        st.error('データベースに接続できませんでした')

# テーブル名をメールアドレスで作成
def create_table_if_not_exists(table_name):
    connection = get_connection()
    cursor = connection.cursor()
    # テーブルが存在するか確認
    cursor.execute(f"SHOW TABLES LIKE '{table_name}'")
    result = cursor.fetchone()

    if not result:
        create_table_query = f"""
        CREATE TABLE `{table_name}` (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(255),
            email VARCHAR(255),
            subject VARCHAR(255),
            date_received DATETIME,
            organization VARCHAR(255),
            status VARCHAR(100),
            tags VARCHAR(255),
            customer_id INT,
            body TEXT,
            summary TEXT
        )"""
        cursor.execute(create_table_query)

    # 接続のクローズ
    cursor.close()
    connection.close()
    

# MYSQLからデータを読み込む関数
def MYSQL_fetch(table_name, choose):
    connection = get_connection()

    if choose == 'ALL':
        cursor = connection.cursor(dictionary=True)
        cursor.execute(f"SELECT * FROM `{table_name}`")
        rows = cursor.fetchall()
    elif choose == 'date_received':
        cursor = connection.cursor(False)
        # 特定のカラムを指定してデータを取得
        cursor.execute(f"SELECT date_received FROM `{table_name}`")
        rows = cursor.fetchall()
        rows = rows[len(rows)-1][0]
    elif choose == 'email':
        cursor = connection.cursor(False)
        # 特定のカラムを指定してデータを取得
        cursor.execute(f"SELECT email FROM `{table_name}`")
        rows = cursor.fetchall()
        rows = list(set(rows))

    # 接続のクローズ
    cursor.close()
    connection.close()

    return rows

# メール本文から顧客情報を抽出する関数
def extract_data(email_body, sender_name, sender_address, subject, raw_date, match_address):
    email_regex = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
    email_match = re.findall(email_regex, email_body)
    
    # 受信日を変換
    parsed_date = parsedate_to_datetime(raw_date)
    jst = timezone(timedelta(hours=9))
    parsed_date = parsed_date.astimezone(jst)
    formatted_date = parsed_date.strftime('%Y-%m-%d %H:%M:%S')
    
    # 要約を生成
    try:
        summary = AI_summary(email_body)
    except Exception as e:
        summary = NLP_summary(email_body)

    # ステータスの取得
    status = "新規"
    for addr in match_address:
        if isinstance(addr, tuple):
            if sender_address == addr[0]:
                status = "既存"
                break
        else:
            if sender_address == addr:
                status = "既存"
                break
    
    if status == "新規":
        match_address.append(sender_address)

    # 仮の組織、タグの設定
    organization = "Unknown Organization"
    tags = "未分類"

    return {
        # 'email': email_match[0] if email_match else None,
        'email': sender_address,
        'sender': sender_name,
        'subject': subject,
        'date_received': formatted_date,
        'organization': organization,
        'status': status,
        'tags': tags,
        'summary': summary
    }

# 抽出したデータをデータベースに保存する関数
def save_to_db(table_name, customer_data):
    connection = get_connection()
    cursor = connection.cursor()

    # SQL挿入クエリ（新しいフィールドを含む）
    query = f"""
        INSERT INTO `{table_name}` (
            name, email, subject, date_received, organization, status, tags, body, summary
        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
    """

    # 顧客データの準備（フィールドがない場合はデフォルト値を使用）
    values = (
        customer_data['sender'],
        customer_data['email'],
        customer_data['subject'],
        customer_data['date_received'],
        customer_data.get('organization', '不明'),
        customer_data['status'],
        customer_data.get('tags', '未分類'),
        customer_data['body'],
        customer_data['summary']
    )

    cursor.execute(query, values)
    connection.commit()
    cursor.close()
    connection.close()

st.title("メールの内容表示および顧客データ抽出アプリ")

cols = st.columns(2)
with cols[0]:
    mail_address = st.text_input('mail_address', placeholder='メールアドレスを入力', label_visibility="hidden")
with cols[1]:
    mail_password = st.text_input('mail_password', placeholder='パスワードを入力', type='password', label_visibility="hidden")

button = st.button("メールを取得")

if 'tabs' not in st.session_state:
    st.session_state.tabs = []

crm_database = []
# ボタンを押すとメールを取得
if button:
    # delete_all_customers()
    try:
        emails = fetch_emails(mail_address, mail_password)
        if mail_address not in st.session_state.tabs:  # 新しいタブのみ追加
            st.session_state.tabs.append(mail_address)
        
        with open('shared_data.txt', 'w') as f:
            f.write('\n'.join(st.session_state.tabs))

        if emails:
            # 各メールを表示し、顧客データを抽出
            try:
                match_address = MYSQL_fetch(mail_address, 'email')
            except Exception as e:
                match_address = []

            for email_data in emails:
                # 顧客データを抽出
                customer_data = extract_data(
                    email_data['body'], email_data['sender_name'], email_data['sender_address'], email_data['subject'], email_data['date'], match_address
                )

                # bodyを追加
                tmp_data = OrderedDict()
                for key, value in customer_data.items():
                    if key == "summary":
                        tmp_data["body"] = email_data['body']
                    tmp_data[key] = value
                customer_data = tmp_data

                # テーブルを作成
                create_table_if_not_exists(mail_address)

                # データベースに保存
                if customer_data['email']:
                    save_to_db(mail_address, customer_data)
                else:
                    st.write("顧客データが不完全です")
        crm_database = MYSQL_fetch(mail_address, 'ALL')
    except EncodingWarning as e:
        st.error('メールデータを取得できませんでした')

if 'datas' not in st.session_state:
    st.session_state.datas = {}

if 'select' not in st.session_state:
    st.session_state.select = {}

if st.session_state.tabs:
    if crm_database:
        st.session_state.datas[mail_address] = crm_database
    tabs = st.tabs(st.session_state.tabs)
    for i, tab in enumerate(tabs):
        with tab:
            col1, col2 = st.columns([2,3])
            # 左側のにメールリストを表示
            with col1:
                st.write("### メール一覧")
                selected_email = None
                address = st.session_state.tabs[i]
                for j in range(len(st.session_state.datas[address])):
                    if st.session_state.datas[address][j]['name']:
                        name = st.session_state.datas[address][j]['name']
                    else:
                        name = ''
                    subject = st.session_state.datas[address][j]['subject']
                    date_received = st.session_state.datas[address][j]['date_received']
                    key = f"{address}_{st.session_state.datas[address][j]['id']}"
                    if st.button(f"{name} \n\n {subject} \n\n {date_received}", use_container_width=True, key=key):
                        st.session_state.select[address] = st.session_state.datas[address][j]

            # 右側のカラムにメール本文を表示
            with col2:
                st.write("### メール本文")
                if address in st.session_state.select:
                    st.write(f"**差出人**: {st.session_state.select[address]['name']} ({st.session_state.select[address]['email']})")
                    st.write(f"**日時**: {st.session_state.select[address]['date_received']}")
                    st.write(f"**件名**: {st.session_state.select[address]['subject']}")
                    st.write(f"**本文**:\n\n{st.session_state.select[address]['body']}")
                else:
                    st.write("メールを選択してください。")
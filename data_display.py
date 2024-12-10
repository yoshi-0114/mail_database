import streamlit as st
import pandas as pd
import mysql.connector
import io
from streamlit_autorefresh import st_autorefresh

st.set_page_config(layout="wide")

# 自動更新を5分（300秒）ごとに設定
st_autorefresh(interval=300 * 1000, limit=100, key="page_autorefresh")

# データベースへの接続を定義
host = ""
user = ""
password = ""
database = ""

# MySQLへの接続設定
def get_connection():
    try:
        connection = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database
        )
        return connection
    except Exception as e:
        st.error('データベースに接続できませんでした')
        return None

# MySQLからデータを読み込む関数
def MYSQL_fetch(table_name):
    connection = get_connection()
    if connection is None:
        return []
    cursor = connection.cursor(dictionary=True)
    cursor.execute(f"SELECT * FROM `{table_name}`")
    rows = cursor.fetchall()
    cursor.close()
    connection.close()
    return rows

# IDが偶数の行に色を付けるスタイリング関数
def highlight_even_rows(row):
    return ['background-color: lightblue' if row['status'] == '新規' else '' for _ in row]

with open('shared_data.txt', 'r') as f:
    mail_address = f.readlines()  # 各行をリストとして読み込む
# 各行の末尾の改行を削除
mail_address = [value.strip() for value in mail_address]

st.title('データベースの内容表示')

if mail_address:
    tabs = st.tabs(mail_address)
    for i, tab in enumerate(tabs):
        with tab:
            # MySQLから取得したデータのDataFrame化とチェックボックス列の追加
            data = MYSQL_fetch(mail_address[i])
            df = pd.DataFrame(data)

            if df.empty:
                st.warning(f"{mail_address[i]}に関連するデータはありません。")
                continue  # データが空の場合は次のタブに進む

            # チェックボックス列の追加
            df.insert(0, '選択', [False] * len(df))

            # マルチセレクトとダウンロードボタンを横並びに配置
            col1, col2 = st.columns([6, 1])

            with col1:
                # マルチセレクトで行を選択
                selected_indices = st.multiselect(
                    "マルチセレクト",
                    df.index,
                    placeholder="ダウンロードしたい行のIDを選択してください",
                    format_func=lambda x: f"{df.loc[x, 'status']} (ID: {df.loc[x, 'id']})",
                    label_visibility="collapsed",
                    key=f'multiselect_{i}'
                )

            # 選択された行をTrueに設定
            df.loc[selected_indices, '選択'] = True

            # ダウンロードボタンを常に表示
            with col2:
                csv_buffer = io.BytesIO()
                df_true = df[df['選択'] == True].drop(columns=['選択']).reset_index(drop=True)
                
                if not df_true.empty:
                    # CSVデータのエンコードとダウンロードボタン
                    df_true.to_csv(csv_buffer, index=False, encoding='shift_jis', errors='ignore')
                    csv_data = csv_buffer.getvalue()
                    
                    st.download_button(
                        label="CSVダウンロード",
                        data=csv_data,
                        file_name="selected_rows.csv",
                        mime="text/csv",
                        key=f'download_button_{i}'
                    )
                else:
                    st.download_button(
                        label="CSVダウンロード",
                        data="",
                        file_name="selected_rows.csv",
                        mime="text/csv",
                        disabled=True,
                        key=f'download_button_disabled_{i}'
                    )

            # スタイルを適用
            styled_df = df.style.apply(highlight_even_rows, axis=1)

            # スタイル付きのDataFrameを表示
            st.dataframe(styled_df, hide_index=True)
else:
    st.error('データベースに保存されていません')
import streamlit as st
import pandas as pd
import os
import time
import json
import re
from datetime import datetime
import PIL.Image
from google import genai
from google.genai import types
from google.oauth2.service_account import Credentials
import gspread

# ==========================================
# 0. ページ設定とCSS
# ==========================================
st.set_page_config(page_title="スマート家計簿ダッシュボード", layout="wide", page_icon="🧾")

# ==========================================
# 0.5. セキュリティロックと状態管理
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# 画像アップローダーをリセットするための魔法の鍵
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

# 前回のアップロード結果を保持するハコ
if "upload_results" not in st.session_state:
    st.session_state["upload_results"] = None

if not st.session_state["authenticated"]:
    st.markdown("<h2 style='text-align: center;'>🔒 秘密の家計簿</h2>", unsafe_allow_html=True)
    pin_input = st.text_input("暗証番号を入力してください", type="password")
    
    if pin_input == st.secrets["APP_PIN"]:
        st.session_state["authenticated"] = True
        st.rerun()
    elif pin_input:
        st.error("❌ 暗証番号が違います")
        
    st.stop() 

st.markdown("""
<style>
    .stApp { background-color: #121212; color: #E0E0E0; }
    div[data-testid="stMetric"] { background-color: #1E1E1E; border: 1px solid #333; border-radius: 12px; padding: 15px; }
    .balance-positive { color: #4CAF50; font-size: 24px; font-weight: bold; text-align: center; margin-top: 10px; }
    .balance-negative { color: #F44336; font-size: 24px; font-weight: bold; text-align: center; margin-top: 10px; }
    .budget-warning { color: #FF9800; font-weight: bold; }
    .budget-danger { color: #F44336; font-weight: bold; }
    .stButton>button { width: 100%; font-weight: bold; height: 50px; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 1. セキュリティ
# ==========================================
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    SHEET_ID = st.secrets["SPREADSHEET_ID"]
    GCP_CREDS_DICT = json.loads(st.secrets["GCP_JSON"])
except Exception as e:
    st.error(f"⚠️ Secrets（秘密の金庫）エラー: {e}")
    st.stop()

# ==========================================
# 2. Googleスプレッドシート連携
# ==========================================
@st.cache_resource(ttl=600)
def init_gspread():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credentials = Credentials.from_service_account_info(GCP_CREDS_DICT, scopes=scopes)
    gc = gspread.authorize(credentials)
    return gc

gc = init_gspread()
sh = gc.open_by_key(SHEET_ID)

def ensure_worksheet(title, headers):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=20)
        ws.append_row(headers)
    return ws

ws_receipts = ensure_worksheet("receipts", ['処理日時', 'レシート日付', '店舗名', '合計金額', '商品名', '金額', '大分類', '小分類'])
ws_income = ensure_worksheet("income", ['年月', '金額'])
ws_fixed = ensure_worksheet("fixed_costs", ['項目名', '金額'])
ws_settings = ensure_worksheet("settings", ['項目', '値'])

@st.cache_data(ttl=60)
def load_receipts():
    data = ws_receipts.get_all_records()
    if not data: return pd.DataFrame(columns=['処理日時', 'レシート日付', '店舗名', '合計金額', '商品名', '金額', '大分類', '小分類'])
    df = pd.DataFrame(data)
    df['レシート日付'] = pd.to_datetime(df['レシート日付'], errors='coerce')
    df['年月'] = df['レシート日付'].dt.strftime('%Y-%m')
    df['金額'] = pd.to_numeric(df['金額'], errors='coerce').fillna(0).astype(int)
    return df

def load_income():
    data = ws_income.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=['年月', '金額'])

def load_fixed_costs():
    data = ws_fixed.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame(columns=['項目名', '金額'])

def load_budget():
    data = ws_settings.get_all_records()
    for row in data:
        if row.get('項目') == 'monthly_budget':
            return int(row.get('値', 100000))
    return 100000

def save_df_to_sheet(ws, df):
    ws.clear()
    ws.update([df.columns.values.tolist()] + df.values.tolist())

raw_df = load_receipts()
income_df = load_income()
fixed_df = load_fixed_costs()
budget = load_budget()

# ==========================================
# 3. AI処理エンジン ＋ 重複チェック
# ==========================================
PROMPT = """
あなたは高精度な家計簿アシスタントです。レシート画像から購入品目を抽出し、JSONで出力してください。
「小計」「合計」「お釣り」「税」は除外。割引はマイナス金額として「その他 > 雑費」に分類。
【カテゴリ指定】
食費、日用品、交通・車両、趣味・娯楽、美容・被服、医療・健康、その他
【出力フォーマット】
{"receipt_date": "YYYY-MM-DD", "store_name": "店舗名", "total_amount": 1500, "items": [{"item_name": "商品名", "price": 200, "main_category": "大分類", "sub_category": "小分類"}]}
"""

def clean_json_string(raw_text):
    text = raw_text.strip()
    if text.startswith("```json"): text = text[7:]
    if text.endswith("```"): text = text[:-3]
    return re.sub(r'[\x00-\x1f\x7f-\x9f]', '', text).strip()

def process_uploaded_files(uploaded_files, current_df):
    client = genai.Client(api_key=API_KEY)
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 【追加機能】すでに登録されているレシートの組み合わせをリストアップ
    existing_receipts = set()
    if not current_df.empty:
        for _, row in current_df.iterrows():
            d_val = row['レシート日付']
            d_str = d_val.strftime('%Y-%m-%d') if pd.notna(d_val) else ""
            existing_receipts.add(f"{d_str}_{row['店舗名']}_{row['合計金額']}")

    results = [] # 処理結果のレポート用
    new_rows = []
    total_files = len(uploaded_files)

    for i, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"処理中 ({i+1}/{total_files}): {uploaded_file.name} をAIが解析しています...")
        try:
            img = PIL.Image.open(uploaded_file)
            response = client.models.generate_content(
                model='gemini-flash-lite-latest',
                contents=[PROMPT, img],
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1)
            )
            
            data = json.loads(clean_json_string(response.text))
            
            # 抽出したデータ
            r_date = data.get('receipt_date', '').replace('/', '-')
            s_name = data.get('store_name', '')
            t_amt = data.get('total_amount', 0)
            
            # 重複チェックの鍵（日付 + 店舗名 + 合計金額）
            dup_key = f"{r_date}_{s_name}_{t_amt}"
            
            if dup_key in existing_receipts:
                # 完全に一致する過去データがあればスキップ
                results.append({"file": uploaded_file.name, "status": "duplicate", "msg": f"重複スキップ: {s_name} ({t_amt}円)"})
            else:
                # 新規レシートの場合
                existing_receipts.add(dup_key) # 同時アップロード同士の重複も防ぐ
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                items = data.get('items', [])
                
                if not items:
                    results.append({"file": uploaded_file.name, "status": "error", "msg": "商品データが読み取れませんでした"})
                else:
                    for item in items:
                        new_rows.append([
                            now, r_date, s_name, t_amt,
                            item.get('item_name', ''), item.get('price', 0), item.get('main_category', ''), item.get('sub_category', '')
                        ])
                    results.append({"file": uploaded_file.name, "status": "success", "msg": f"保存完了: {s_name} ({t_amt}円)"})
                    
        except Exception as e:
            results.append({"file": uploaded_file.name, "status": "error", "msg": f"読取エラーが発生しました"})
        
        progress_bar.progress((i + 1) / total_files)
        if i < total_files - 1:
            status_text.text("AIの思考を整理中... (数秒お待ちください)")
            time.sleep(15)

    if new_rows:
        ws_receipts.append_rows(new_rows)

    # 処理結果を保存
    st.session_state["upload_results"] = results
    status_text.text("✅ 処理が完了しました。")
    
    # 1つでも「error」があればFalseを返す（duplicateはエラー扱いしない）
    has_error = any(r['status'] == 'error' for r in results)
    return not has_error

# ==========================================
# 4. メイン画面UI
# ==========================================
st.markdown("<h1 style='text-align: center;'>📱 スマート家計簿</h1>", unsafe_allow_html=True)

tab_input, tab_monthly, tab_trend, tab_settings = st.tabs([
    "📤 レシート入力 (Do)", 
    "📊 今月の収支 (Check)", 
    "📈 月別トレンド (Act)", 
    "⚙️ 設定・データ編集"
])

# ------------------------------------------
# 📤 タブ1：レシート入力
# ------------------------------------------
with tab_input:
    # 【追加機能】前回の処理結果レポートを表示
    if st.session_state.get("upload_results"):
        st.markdown("### 📝 前回の処理結果")
        for res in st.session_state["upload_results"]:
            if res['status'] == 'success':
                st.success(f"✅ {res['file']} -> {res['msg']}")
            elif res['status'] == 'duplicate':
                st.warning(f"⏩ {res['file']} -> {res['msg']} (すでに登録済みのためスキップしました)")
            else:
                st.error(f"❌ {res['file']} -> {res['msg']}")
        
        if st.button("結果の表示を消す"):
            st.session_state["upload_results"] = None
            st.rerun()
        st.markdown("---")

    st.subheader("📸 レシートの追加")
    st.markdown("スマホの場合は「ファイルを参照」を押すと、**その場でカメラを起動**できます。複数枚の同時アップロードも可能です。")
    
    uploaded_files = st.file_uploader(
        "レシート画像をアップロード", 
        type=['png', 'jpg', 'jpeg', 'webp'], 
        accept_multiple_files=True, 
        key=f"uploader_{st.session_state['uploader_key']}"
    )
    
    if uploaded_files:
        if st.button("🚀 このレシートをAIで解析する", type="primary"):
            # 現在のデータを渡して重複チェックさせる
            is_all_success_or_dup = process_uploaded_files(uploaded_files, raw_df)
            
            st.cache_data.clear()
            
            if is_all_success_or_dup:
                # 成功＆重複スキップのみの場合は、アップローダーを綺麗にする
                time.sleep(1.5)
                st.session_state["uploader_key"] += 1
                st.rerun()
            else:
                # エラー（読み取り不能など）があった場合は画像を残す
                st.warning("⚠️ 一部の画像でエラーが発生したため、画像を残しています。上の結果レポートを確認してください。")

# ------------------------------------------
# 📊 タブ2：今月の収支
# ------------------------------------------
with tab_monthly:
    if raw_df.empty:
        st.info("データがありません。「レシート入力」タブから画像をアップロードしてください。")
    else:
        available_months = sorted(raw_df['年月'].dropna().unique(), reverse=True)
        income_months = income_df['年月'].tolist() if not income_df.empty else []
        all_months = sorted(list(set(available_months + income_months)), reverse=True)
        
        selected_month = st.selectbox("📅 表示月", all_months if all_months else [datetime.today().strftime('%Y-%m')])
        df_m = raw_df[raw_df['年月'] == selected_month]
        
        # 収支計算
        month_income = income_df[income_df['年月'] == selected_month]['金額'].sum() if not income_df.empty else 0
        total_fixed = fixed_df['金額'].sum() if not fixed_df.empty else 0
        total_variable = df_m['金額'].sum()
        balance = month_income - total_fixed - total_variable
        
        st.subheader("👛 全体収支")
        c1, c2, c3 = st.columns(3)
        c1.metric("① 収入", f"{month_income:,} 円")
        c2.metric("② 固定費", f"{total_fixed:,} 円")
        c3.metric("③ 変動費（今月使った額）", f"{total_variable:,} 円")
        
        balance_text = f"<div class='balance-positive'>④ 最終的な手残り: +{balance:,} 円</div>" if balance >= 0 else f"<div class='balance-negative'>④ 最終的な手残り: {balance:,} 円</div>"
        st.markdown(balance_text, unsafe_allow_html=True)
        st.markdown("---")
        
        st.subheader("🚦 今月の予算消化ペース")
        remaining_budget = budget - total_variable
        spend_ratio = min(total_variable / budget if budget > 0 else 1.0, 1.0)
        
        st.progress(spend_ratio)
        if spend_ratio < 0.7:
            st.success(f"✅ 余裕があります。上限まであと **{remaining_budget:,} 円** 使えます。")
        elif spend_ratio < 1.0:
            st.warning(f"⚠️ 予算に近づいています。上限まであと **{remaining_budget:,} 円** です。")
        else:
            st.error(f"🚨 予算オーバーです！ 上限を **{-remaining_budget:,} 円** 超過しています。")

        today = datetime.today()
        if selected_month == today.strftime('%Y-%m'):
            days_passed = today.day
            import calendar
            days_in_month = calendar.monthrange(today.year, today.month)[1]
            days_left = days_in_month - days_passed + 1 
        else:
            days_passed = 30; days_left = 0
            
        daily_avg = total_variable // days_passed if days_passed > 0 else 0
        daily_allowance = remaining_budget // days_left if days_left > 0 else 0
        num_transactions = len(df_m[['レシート日付', '店舗名']].drop_duplicates())
        avg_per_transaction = total_variable // num_transactions if num_transactions > 0 else 0

        with st.expander("🔍 変動費の詳細分析・購入履歴を見る"):
            k1, k2 = st.columns(2)
            k1.metric("今日から1日あたり使える額", f"{daily_allowance:,} 円 / 日" if days_left > 0 and remaining_budget > 0 else "-")
            k2.metric("これまでの1会計あたり平均", f"{avg_per_transaction:,} 円 / 回")
            
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.write("**大分類別の支出**")
                st.bar_chart(df_m.groupby('大分類')['金額'].sum())
            with col_g2:
                st.write("**小分類ランキング**")
                st.dataframe(df_m.groupby(['大分類', '小分類'])['金額'].sum().sort_values(ascending=False), width='stretch')

            display_df = df_m[['レシート日付', '店舗名', '商品名', '金額', '大分類']].copy()
            if not display_df.empty:
                display_df['レシート日付'] = display_df['レシート日付'].dt.strftime('%Y-%m-%d')
            st.dataframe(display_df.sort_values('レシート日付', ascending=False), hide_index=True, width='stretch')

# ------------------------------------------
# 📈 タブ3：月別トレンド
# ------------------------------------------
with tab_trend:
    st.subheader("📈 月別の支出推移と傾向（変動費）")
    if raw_df.empty:
        st.info("データがありません。")
    else:
        st.markdown("**■ 変動費 総額の推移**")
        monthly_trend = raw_df.groupby('年月')['金額'].sum()
        st.line_chart(monthly_trend)
        
        st.markdown("**■ カテゴリ別の推移**")
        pivot_main = raw_df.pivot_table(index='年月', columns='大分類', values='金額', aggfunc='sum', fill_value=0)
        st.bar_chart(pivot_main)

# ------------------------------------------
# ⚙️ タブ4：設定・データ編集
# ------------------------------------------
with tab_settings:
    st.subheader("バックヤード設定 (クラウド同期)")
    st.markdown("ここで編集した内容は、即座にGoogleスプレッドシートに保存されます。")
    
    with st.expander("1️⃣ システム初期設定 (予算)"):
        new_budget = st.number_input("毎月の費用上限額（変動費の目標/円）", value=budget, step=10000)
        if st.button("💾 予算設定を保存"):
            df_settings = pd.DataFrame([{'項目': 'monthly_budget', '値': new_budget}])
            save_df_to_sheet(ws_settings, df_settings)
            st.success("保存しました！")
            time.sleep(1); st.rerun()
            
    with st.expander("2️⃣ 月別収入の登録"):
        edited_inc_df = st.data_editor(income_df if not income_df.empty else pd.DataFrame([{"年月": datetime.today().strftime('%Y-%m'), "金額": 0}]), num_rows="dynamic", width='stretch', key="inc")
        if st.button("💾 収入データを保存"):
            save_df_to_sheet(ws_income, edited_inc_df.dropna(subset=['年月']))
            st.success("保存しました！")
            time.sleep(1); st.rerun()

    with st.expander("3️⃣ 固定費の登録"):
        edited_fc_df = st.data_editor(fixed_df if not fixed_df.empty else pd.DataFrame([{"項目名": "", "金額": 0}]), num_rows="dynamic", width='stretch', key="fc")
        if st.button("💾 固定費を保存"):
            save_df_to_sheet(ws_fixed, edited_fc_df.dropna(subset=['項目名']))
            st.success("保存しました！")
            time.sleep(1); st.rerun()

    with st.expander("4️⃣ 生データ（変動費）の編集・削除"):
        if not raw_df.empty:
            edited_raw_df = st.data_editor(raw_df, num_rows="dynamic", width='stretch', key="raw")
            if st.button("💾 上記の内容で生データを上書き保存"):
                columns_to_save = ['処理日時', 'レシート日付', '店舗名', '合計金額', '商品名', '金額', '大分類', '小分類']
                save_df_to_sheet(ws_receipts, edited_raw_df[columns_to_save])
                st.cache_data.clear()
                st.success("保存しました！")
                time.sleep(1); st.rerun()
        else:
            st.info("まだレシートデータがありません。")

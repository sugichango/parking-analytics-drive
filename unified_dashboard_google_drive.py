import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import io
import json
import gzip
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# 0. Google Drive 連携設定 (追加部分)
# ==========================================

@st.cache_resource
def get_drive_service():
    """Streamlit Secretsから認証情報を取得してGoogle Driveサービスを構築"""
    try:
        info = json.loads(st.secrets["gcp_service_account"])
        credentials = service_account.Credentials.from_service_account_info(info)
        service = build('drive', 'v3', credentials=credentials)
        return service
    except Exception as e:
        if "gcp_service_account" not in st.secrets:
            st.error("Streamlit Secrets に 'gcp_service_account' が設定されていません。")
        else:
            st.error(f"Google Drive連携エラー: {e}")
        return None

def download_file_from_drive(file_id):
    """ファイルIDを指定してバイナリデータをメモリ上にダウンロード"""
    service = get_drive_service()
    if not service: return None
    try:
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except Exception: return None

def search_files_in_drive(query):
    """Googleドライブ内を検索 (共有ドライブ等にも対応)"""
    service = get_drive_service()
    if not service: return []
    try:
        results = service.files().list(
            q=query, 
            fields="files(id, name, parents, modifiedTime)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()

        return results.get('files', [])
    except Exception: return []

# ==========================================
# 0. 共通：初期設定と認証機能 (原本を忠実に再現)
# ==========================================
st.set_page_config(
    page_title="駐車場 統合アナリティクス",
    page_icon="🅿️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- 全体デザイン＆CSS設定 (原本を忠実に再現) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    
    html, body, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {
        background-color: #0E1117 !important;
        color: #FFFFFF !important;
        font-family: 'Inter', sans-serif !important;
    }
    
    [data-testid="stSidebar"] {
        background-color: #161B22 !important;
    }
    
    .login-container {
        max-width: 420px;
        margin: 80px auto 0 auto;
        background: rgba(255,255,255,0.05);
        border: 1px solid rgba(0,255,255,0.2);
        border-radius: 18px;
        padding: 48px 40px 40px 40px;
        box-shadow: 0 8px 40px rgba(0,0,0,0.6);
    }
    .login-title {
        text-align: center;
        font-size: 28px;
        font-weight: 800;
        color: #00FFFF;
        text-shadow: 0 0 16px rgba(0,255,255,0.5);
        margin-bottom: 6px;
    }

    /* KPIカード領域 */
    div[data-testid="stMetric"] {
        background-color: rgba(255, 255, 255, 0.12) !important;
        border: 1px solid rgba(0, 255, 255, 0.25) !important;
        padding: 24px !important;
        border-radius: 14px !important;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5) !important;
    }
    div[data-testid="stMetricValue"] > div {
        font-size: 38px !important;
        font-weight: 800 !important;
        color: #00FFFF !important;
        text-shadow: 0 0 12px rgba(0, 255, 255, 0.4) !important;
    }
    [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] p {
        color: #FFFFFF !important;
        font-weight: 900 !important;
        font-size: 16px !important;
        opacity: 1.0 !important;
        text-shadow: 1px 1px 3px rgba(0,0,0,1.0) !important;
    }
    /* ヘルプアイコン (?) の抜本的な視認性修正 */
    [data-testid="stMetricLabel"] button {
        background-color: #FFFF00 !important;
        border: 2px solid #000000 !important;
        border-radius: 50% !important;
        width: 32px !important;
        height: 32px !important;
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        margin-left: 10px !important;
        opacity: 1.0 !important;
        box-shadow: 0 0 12px rgba(255, 255, 0, 0.7) !important;
        position: relative !important;
    }
    [data-testid="stMetricLabel"] button svg { display: none !important; }
    [data-testid="stMetricLabel"] button::before {
        content: "?" !important;
        color: #000000 !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        font-family: 'Arial Black', sans-serif !important;
        display: block !important;
        line-height: 1 !important;
    }
    /* ラベル・チェックボックス・全テキストを純白・極太に */
    label[data-testid="stWidgetLabel"], .stSelectbox label, .stRadio label, .stMultiSelect label, .stCheckbox p, .stCheckbox span, .stCheckbox label {
        color: #FFFFFF !important;
        font-weight: 900 !important;
        font-size: 16px !important;
        opacity: 1.0 !important;
        text-shadow: 1px 1px 3px rgba(0,0,0,0.8) !important;
    }
</style>
""", unsafe_allow_html=True)

def show_login_page():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if st.session_state["authenticated"]:
        return True

    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-title">🅿️ 駐車場利用データ分析<br>ダッシュボード</div>', unsafe_allow_html=True)
    
    email = st.text_input("メールアドレス（例: test@tutc.or.jp）", key="login_email")
    password = st.text_input("パスワード", type="password", key="login_pass")

    if st.button("ログイン", key="login_btn"):
        correct_password = st.secrets.get("app_password", st.secrets.get("LOGIN_PASSWORD", "tutc_secure_login"))
        if not email.strip().lower().endswith("@tutc.or.jp"):
            st.error("許可されていないドメインです。")
        elif password == correct_password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")
    st.markdown('</div>', unsafe_allow_html=True)
    return False

# ==========================================
# メイン画面とメニュー切替
# ==========================================
if not show_login_page():
    st.stop()

# 共有フォルダID（Secretsから取得）
DRIVE_FOLDER_ID = st.secrets.get("google_drive_folder_id", "")

with st.sidebar:
    st.markdown("---")
    st.subheader("🛠️ ツール切替")
    mode = st.radio("表示メニュー", ["① 一般利用台数推移分析", "② 24時間稼働状況分析"], key="sys_mode")
    
    # 使い方マニュアルのダウンロード (Drive版)
    if st.button("📖 使い方マニュアルを検索", key="search_manual"):
        q = f"name = 'README_user.txt' and trashed = false"
        manual_files = search_files_in_drive(q)
        if manual_files:
            fh = download_file_from_drive(manual_files[0]['id'])
            if fh:
                st.download_button(
                    label="📥 マニュアルを保存",
                    data=fh.getvalue(),
                    file_name="駐車場データダッシュボード_使い方マニュアル.txt",
                    mime="text/plain"
                )
        else:
            st.warning("マニュアルファイル(README_user.txt)が見つかりませんでした。")

    if st.button("ログアウト", key="logout_act"):
        st.session_state["authenticated"] = False
        st.rerun()
    
    st.markdown("---")
    if st.button("🔄 キャッシュをクリアして更新", key="clear_cache_btn"):
        st.cache_data.clear()
        st.success("キャッシュをクリアしました。最新データを再読み込みします。")
        st.rerun()
    st.markdown("---")


# ==========================================
# ① 一般利用台数推移分析 (原本ロジック完全再現)
# ==========================================
if "①" in mode:
    st.title("📊 ① 一般利用台数推移分析")

    @st.cache_data(show_spinner=True)
    def load_data_dashboard1_drive(file_name, modified_time):
        """Google DriveからCSVを取得し原本と同じロジックで加工（modified_timeでキャッシュ自動更新）"""

        q = f"name = '{file_name}' and trashed = false"
        files = search_files_in_drive(q)
        if not files:
            st.error(f"ファイルが見つかりません: {file_name}")
            return None
        fh = download_file_from_drive(files[0]['id'])
        if not fh:
            st.error(f"ファイルのダウンロードに失敗しました: {file_name}")
            return None
        
        use_cols = ['ParkingArea', 'OnTime', 'Cash', 'Discount1', 'Discount2', 'Discount3', 'Discount4', 'Discount5', 'Discount6', 'Discount7']
        dtypes = {'ParkingArea': 'Int16', 'Cash': 'Int32', 'Discount1': 'Int16', 'Discount2': 'Int16', 'Discount3': 'Int16', 'Discount4': 'Int16', 'Discount5': 'Int16', 'Discount6': 'Int16', 'Discount7': 'Int16'}
        
        try:
            if file_name.endswith(".gz"):
                with gzip.open(fh, 'rt', encoding='utf-8') as f:
                    df = pd.read_csv(f, usecols=lambda c: c in use_cols, dtype=dtypes, parse_dates=['OnTime'] if 'OnTime' in use_cols else False)
            else:
                df = pd.read_csv(fh, usecols=lambda c: c in use_cols, dtype=dtypes, parse_dates=['OnTime'] if 'OnTime' in use_cols else False)
        except Exception:
            fh.seek(0)
            try:
                if file_name.endswith(".gz"):
                    with gzip.open(fh, 'rt', encoding='cp932') as f:
                        df = pd.read_csv(f, usecols=lambda c: c in use_cols, dtype=dtypes, parse_dates=['OnTime'] if 'OnTime' in use_cols else False)
                else:
                    df = pd.read_csv(fh, encoding='cp932', usecols=lambda c: c in use_cols, dtype=dtypes, parse_dates=['OnTime'] if 'OnTime' in use_cols else False)
            except Exception: return None

        if df is not None:
            parking_areas = {440: "南１", 441: "南２", 442: "南３", 443: "南４", 444: "北１", 445: "北２", 446: "北３"}
            if 'ParkingArea' in df.columns: df['ParkingAreaName'] = df['ParkingArea'].map(parking_areas).fillna('不明')
            discount_cols = [f'Discount{i}' for i in range(1, 8)]
            active_disc_cols = [c for c in discount_cols if c in df.columns]
            for c in active_disc_cols: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype('Int64')
            mask_rb = pd.Series(False, index=df.index); mask_ticket = pd.Series(False, index=df.index); mask_any = pd.Series(False, index=df.index)
            rb_codes = [11, 12, 13, 14, 15, 43, 44]; ticket_codes = [30, 31, 32, 33, 34, 35]
            for c in active_disc_cols:
                mask_rb = mask_rb | df[c].isin(rb_codes); mask_ticket = mask_ticket | df[c].isin(ticket_codes); mask_any = mask_any | (df[c] > 0)
            df['PaymentType'] = 'その他'; df.loc[~mask_any, 'PaymentType'] = '現金'; df.loc[mask_ticket & ~mask_rb, 'PaymentType'] = '回数券'; df.loc[mask_rb, 'PaymentType'] = 'RB'
            if 'OnTime' in df.columns:
                df['OnTime'] = pd.to_datetime(df['OnTime'], errors='coerce'); df['Month'] = df['OnTime'].dt.to_period('M').astype(str); df['DayOfWeek'] = df['OnTime'].dt.dayofweek
                df['is_holiday'] = df['DayOfWeek'].isin([5, 6]).map({True: '休日', False: '平日'})
        return df


    @st.cache_data(show_spinner=False)
    def load_comparison_monthly(_file_keys):
        """全年度ファイルを(fiscal_year, ParkingAreaName, Month)で月次集計して返す。
        集計済みの小DataFrameだけキャッシュするため高速。"""
        import re as _re2
        rows = []
        for _fname, _mtime in _file_keys:
            _m = _re2.search(r'(202[0-9])', _fname)
            if not _m: continue
            _fy = int(_m.group(1))
            _df = load_data_dashboard1_drive(_fname, _mtime)
            if _df is None or _df.empty: continue
            if 'Month' not in _df.columns or 'ParkingAreaName' not in _df.columns: continue
            _cash_col = 'Cash' if 'Cash' in _df.columns else None
            if _cash_col:
                _agg = _df.groupby(['ParkingAreaName', 'Month']).agg(
                    利用台数=('OnTime', 'count'), 現金収入=(_cash_col, 'sum')
                ).reset_index()
            else:
                _agg = _df.groupby(['ParkingAreaName', 'Month']).agg(
                    利用台数=('OnTime', 'count')
                ).reset_index()
                _agg['現金収入'] = 0
            _agg['fiscal_year'] = _fy
            rows.append(_agg)
        if not rows:
            return pd.DataFrame()
        return pd.concat(rows, ignore_index=True)

    # --- 利用可能な年度の特定ロジック (最強版) ---
    q_years = "name contains 'updated_integrated_data' and trashed = false"
    drive_files = search_files_in_drive(q_years)
    
    # 候補となる年度リスト
    candidate_years = []
    
    # 1. ファイル名から抽出を試みる
    if drive_files:
        import re
        for f in drive_files:
            m = re.search(r'202[0-9]', f['name'])
            if m: candidate_years.append(m.group())

    # ひとまずファイル名ベースの年度があれば表示
    available_years = sorted(list(set(candidate_years)), reverse=True)
    
    # サイドバーに年度選択を配置
    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 年度選択")
    
    # もしファイル名に年度がなければ、「全年度(統合)」として一度読み込み、中身から抽出する
    if not available_years:
        available_years = ["全年度(統合)"]
        selected_year_d1 = st.sidebar.selectbox("分析対象年度 (一般利用)", available_years, index=0, key="d1_year_select")
        TARGET_CSV = "updated_integrated_data.csv.gz"
        modified_time = drive_files[0].get('modifiedTime', '') if drive_files else ''
    else:
        selected_year_d1 = st.sidebar.selectbox("分析対象年度 (一般利用)", available_years, index=0, key="d1_year_select")
        # 選択された年度に対応するファイル名を作成（または検索）
        target_f = [f for f in drive_files if selected_year_d1 in f['name'] and f['name'].endswith(".csv.gz")]
        if target_f:
            TARGET_CSV = target_f[0]['name']
            modified_time = target_f[0].get('modifiedTime', '')
        else:
            TARGET_CSV = "updated_integrated_data.csv.gz"
            modified_time = ''

    # データの読み込み
    with st.spinner(f"{TARGET_CSV} を読み込み中..."):
        df_d1 = load_data_dashboard1_drive(TARGET_CSV, modified_time)
    
    # 2. データの中身からさらに年度を抽出（ファイル名に年度がなかった場合の補完）
    if df_d1 is not None and not df_d1.empty:
        # 日付列を探す
        date_col = None
        for col in ['OnTime', 'ontime', '日付', 'Datetime']:
            if col in df_d1.columns:
                date_col = col; break
        
        if date_col:
            df_d1[date_col] = pd.to_datetime(df_d1[date_col], errors='coerce')
            extracted_years = sorted(df_d1[date_col].dt.year.dropna().unique().astype(int).astype(str), reverse=True)
            
            # もし「全年度(統合)」モードで、中身に複数年度あった場合、再度選択させる
            if selected_year_d1 == "全年度(統合)" and len(extracted_years) > 1:
                selected_year_d1 = st.sidebar.selectbox("分析対象年度を選択してください", extracted_years, index=0, key="d1_year_reselect")
                df_d1 = df_d1[df_d1[date_col].dt.year.astype(str) == selected_year_d1]
            elif selected_year_d1 == "全年度(統合)" and len(extracted_years) == 1:
                selected_year_d1 = extracted_years[0]
        
    # --- サイドバー診断情報 (非表示の展開メニュー) ---
    with st.sidebar.expander("🔍 診断情報 (ボタンが出ない場合用)"):
        st.write(f"読み込みファイル: {TARGET_CSV}")
        if drive_files: st.write(f"Drive上のファイル数: {len(drive_files)}")
        if df_d1 is not None: st.write(f"データ行数: {len(df_d1)}")
        if df_d1 is not None: st.write(f"カラム一覧: {df_d1.columns.tolist()}")
        df_d1 = load_data_dashboard1_drive(TARGET_CSV, modified_time)
    
    if df_d1 is None:
        st.error(f"データの読み込みに失敗しました。ファイル名: {TARGET_CSV}")
    elif df_d1.empty:
        st.warning(f"読み込まれたデータが0件です。ファイル名: {TARGET_CSV}")
    else:
        st.success(f"データを読み込みました: {len(df_d1):,} 件")
        with st.expander("🛠️ デバッグ：データの中身を確認"):
            st.write("項目名:", df_d1.columns.tolist())
            st.write("データ見本:", df_d1.head())
            st.write("項目名:", df_d1.columns.tolist())
            st.write("データ見本:", df_d1.head())
            if 'OnTime' in df_d1.columns:
                st.write("日付の範囲:", df_d1['OnTime'].min(), " ～ ", df_d1['OnTime'].max())

    if df_d1 is not None:
        st.sidebar.header("🔍 フィルター設定")
        PARKING_ORDER = ["南１", "南２", "南３", "南４", "北１", "北２", "北３"]
        PARKING_COLORS = {"南１": "#00FFFF", "南２": "#FF00FF", "南３": "#39FF14", "南４": "#FFFF00", "北１": "#FF4500", "北２": "#9D00FF", "北３": "#1E90FF"}
        
        available_areas = ["全駐車場"]
        if 'ParkingAreaName' in df_d1.columns:
            areas = [p for p in PARKING_ORDER if p in df_d1['ParkingAreaName'].unique()]
            others = sorted([p for p in df_d1['ParkingAreaName'].unique() if p not in PARKING_ORDER and p != '不明' and p != '全駐車場'])
            available_areas.extend(areas + others)
        selected_area = st.sidebar.selectbox("駐車場名", available_areas, index=0)
        selected_day_type = st.sidebar.selectbox("平日/休日", ["すべて", "平日", "休日"], index=0)
        
        available_months = ["通年"]
        if 'Month' in df_d1.columns:
            months = sorted(df_d1[df_d1['Month'] != 'NaT']['Month'].unique().tolist())
            available_months.extend(months)
        selected_month = st.sidebar.selectbox("対象月", available_months, index=0)

        filtered_df = df_d1.copy()
        if selected_area != "全駐車場": filtered_df = filtered_df[filtered_df['ParkingAreaName'] == selected_area]
        if selected_day_type != "すべて": filtered_df = filtered_df[filtered_df['is_holiday'] == selected_day_type]
        if selected_month != "通年": filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        st.markdown(f"**現在の絞り込み**: 駐車場=`{selected_area}` | 曜日=`{selected_day_type}` | 月=`{selected_month}` (対象データ: {len(filtered_df):,} 件)")
        
        st.markdown('''<div style="border: 2px solid #00FFFF; border-radius: 8px; padding: 12px; margin-top: 20px; margin-bottom: 5px; background: rgba(0, 255, 255, 0.05);"><b style="color: #00FFFF; font-size: 16px;">🔲 内訳表示オプション</b></div>''', unsafe_allow_html=True)
        show_by_payment_type = st.checkbox("👉 支払い種別（現金・RB・回数券）で内訳を表示する", value=False)
        
        if not filtered_df.empty and 'OnTime' in filtered_df.columns:
            st.subheader("📈 利用台数推移 (月別または日別)")
            if 'Cash' in filtered_df.columns:
                filtered_df['Cash'] = pd.to_numeric(filtered_df['Cash'], errors='coerce').fillna(0)
            else:
                filtered_df['Cash'] = 0

            if selected_month == "通年": x_col, x_title = 'Month', '年月'
            else: filtered_df['Date'] = filtered_df['OnTime'].dt.date.astype(str); x_col, x_title = 'Date', '日付'

            if show_by_payment_type and 'PaymentType' in filtered_df.columns:
                bar_counts = filtered_df.groupby([x_col, 'PaymentType']).size().reset_index(name='利用台数')
                mapping = {'現金': '現金（現金のみ）', '回数券': '回数券（回数券のみ、回数券+現金）', 'RB': 'RB（RBのみ、RB+回数券、RB+現金、RB+回数券+現金）', 'その他': 'その他'}
                bar_counts['PaymentTypeLegend'] = bar_counts['PaymentType'].map(mapping).fillna(bar_counts['PaymentType'])
                color_col = 'PaymentTypeLegend'
            elif selected_area == "全駐車場" and 'ParkingAreaName' in filtered_df.columns:
                bar_counts = filtered_df.groupby([x_col, 'ParkingAreaName']).size().reset_index(name='利用台数')
                color_col = 'ParkingAreaName'
            else:
                bar_counts = filtered_df.groupby(x_col).size().reset_index(name='利用台数')
                color_col = None

            line_counts = filtered_df.groupby(x_col).agg({'Cash': 'sum'}).rename(columns={'Cash': '現金収入'}).reset_index()
            total_counts = bar_counts.groupby(x_col)['利用台数'].sum().reset_index(name='合計台数')
            
            if color_col:
                bar_counts = pd.merge(bar_counts, total_counts, on=x_col)
                bar_counts['割合'] = (bar_counts['利用台数'] / bar_counts['合計台数'] * 100).round(1)
                bar_counts['text'] = bar_counts.apply(lambda row: f"{row['割合']}%" if row['割合'] > 0 else "", axis=1)
            else: bar_counts['text'] = bar_counts['利用台数'].astype(str)

            neon_colors = ['#00FFFF', '#FF00FF', '#39FF14', '#FFEA00', '#FF003C', '#9D00FF', '#00F0FF']
            common_layout = dict(
                font=dict(family="sans-serif", color="#E0E0E0"),
                plot_bgcolor="rgba(17, 17, 17, 1)",
                paper_bgcolor="rgba(17, 17, 17, 1)",
                margin={'l': 30, 'r': 30, 't': 50, 'b': 30}
            )

            fig_bar = make_subplots(specs=[[{"secondary_y": True}]])
            if color_col:
                if color_col == 'ParkingAreaName':
                    plot_cats = [p for p in PARKING_ORDER if p in bar_counts[color_col].unique()]
                    plot_cats.extend([c for c in sorted(bar_counts[color_col].unique()) if c not in PARKING_ORDER])
                    for cat in plot_cats:
                        d = bar_counts[bar_counts[color_col] == cat]
                        fig_bar.add_trace(go.Bar(x=d[x_col], y=d['利用台数'], name=str(cat), text=d['text'], textposition='inside', insidetextanchor='middle', marker_color=PARKING_COLORS.get(cat, "#FFFFFF")), secondary_y=False)
                else:
                    for idx, cat in enumerate(sorted(bar_counts[color_col].unique())):
                        d = bar_counts[bar_counts[color_col] == cat]
                        fig_bar.add_trace(go.Bar(x=d[x_col], y=d['利用台数'], name=str(cat), text=d['text'], textposition='inside', insidetextanchor='middle', marker_color=neon_colors[idx % len(neon_colors)]), secondary_y=False)
                fig_bar.update_layout(barmode='stack')
                for i, row in total_counts.iterrows():
                    fig_bar.add_annotation(x=row[x_col], y=row['合計台数'], text=str(row['合計台数']), showarrow=False, yshift=10, font=dict(color="#00FFFF", size=13, family="sans-serif"), yref="y")
            else:
                fig_bar.add_trace(go.Bar(x=bar_counts[x_col], y=bar_counts['利用台数'], name="利用台数", marker_color=neon_colors[0], text=bar_counts['text'], textposition='auto', textfont=dict(color="black")), secondary_y=False)
            
            line_color = '#FFFFFF'
            fig_bar.add_trace(go.Scatter(x=line_counts[x_col], y=line_counts['現金収入'], name="現金収入(全体)", mode='lines+markers', line={'color': line_color, 'width': 3}), secondary_y=True)
            
            fig_bar.update_layout(**common_layout, title=dict(text=f"{x_title} 利用台数と現金収入推移", font=dict(size=18, color="#00FFFF")), xaxis_title=x_title, legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(color="#E0E0E0"), traceorder="normal"))
            fig_bar.update_xaxes(showgrid=True, gridcolor='rgba(255,255,255,0.1)', type='category', categoryorder='category ascending')
            fig_bar.update_yaxes(title_text="利用台数（台）", secondary_y=False, rangemode='tozero', showgrid=True, gridcolor='rgba(255,255,255,0.1)')
            fig_bar.update_yaxes(title_text="現金収入（円）", secondary_y=True, rangemode='tozero', showgrid=False)
            
            st.plotly_chart(fig_bar, use_container_width=True)
            st.markdown("---")
            
            if 'ParkingAreaName' in filtered_df.columns:
                area_counts = filtered_df.groupby('ParkingAreaName').agg({'OnTime': 'size', 'Cash': 'sum'}).rename(columns={'OnTime': '利用台数', 'Cash': '現金収入'}).reset_index()
                total_parked = area_counts['利用台数'].sum()
                total_cash = area_counts['現金収入'].sum()
                parking_colors = {area: PARKING_COLORS.get(area, "#FFFFFF") for area in filtered_df['ParkingAreaName'].unique()}
                
                if show_by_payment_type:
                    st.subheader("🍩 駐車場別 利用内訳（サンバースト図）")
                    st.write("※ グラフの要素をクリックすると、その階層を拡大（ドリルダウン）できます。中央をクリックすると元に戻ります。")
                    if 'PaymentType' in filtered_df.columns:
                        agg_df = filtered_df.groupby(['ParkingAreaName', 'PaymentType']).agg({'OnTime': 'size', 'Cash': 'sum'}).rename(columns={'OnTime': '利用台数', 'Cash': '現金収入'}).reset_index()
                        fig_chart = px.sunburst(agg_df[agg_df['利用台数']>0], path=['ParkingAreaName', 'PaymentType'], values='利用台数', color='ParkingAreaName', color_discrete_map=parking_colors)
                        fig_chart.update_traces(texttemplate='%{label}<br>%{value}台<br>%{percentRoot}', hovertemplate='%{label}<br>利用台数: %{value}台<br>割合: %{percentRoot}', insidetextorientation='radial')
                        fig_chart.update_layout(**common_layout, height=600)
                        st.plotly_chart(fig_chart, use_container_width=True)
                else:
                    st.subheader("🍩 駐車場別 利用割合 ＆ 現金収入割合")
                    fig_chart = go.Figure()
                    inner_domain, middle_domain = [0.15, 0.85], [0.0, 1.0]
                    inner_hole, middle_hole = 0.55, 0.8
                    fig_chart.add_trace(go.Pie(labels=area_counts['ParkingAreaName'], values=area_counts['利用台数'], name="利用台数", hole=inner_hole, domain={'x': inner_domain, 'y': inner_domain}, hoverinfo='label+value+percent+name', textinfo='label+value+percent', textposition='inside', direction='clockwise', sort=False, marker=dict(colors=[parking_colors.get(label, '#FFFFFF') for label in area_counts['ParkingAreaName']]), insidetextfont=dict(color="black")))
                    fig_chart.add_trace(go.Pie(labels=area_counts['ParkingAreaName'], values=area_counts['現金収入'], name="現金収入", hole=middle_hole, domain={'x': middle_domain, 'y': middle_domain}, hoverinfo='label+value+percent+name', textinfo='value+percent', textposition='inside', direction='clockwise', sort=False, marker=dict(colors=[parking_colors.get(label, '#FFFFFF') for label in area_counts['ParkingAreaName']]), insidetextfont=dict(color="black")))
                    fig_chart.update_layout(**common_layout, title=dict(text="内側: 利用台数 / 外側: 現金収入", font=dict(color="#00FFFF")), annotations=[{"text": f"総台数<br><b style='font_size:20px;'>{total_parked:,}</b><br>台<br><br>総現金<br><b style='font_size:16px;'>{int(total_cash):,}</b><br>円", "x": 0.5, "y": 0.5, "showarrow": False, "font": dict(color="#00FFFF")}], showlegend=True, legend=dict(font=dict(color="#E0E0E0")))
                    st.plotly_chart(fig_chart, use_container_width=True)
                st.markdown("---")

        # ==========================================
        # 年度間比較分析セクション
        # ==========================================
        st.markdown("---")
        show_cmp = st.checkbox("📊 年度間比較分析を表示する（クリックでデータを読み込みます）", value=False, key="show_cmp")
        if not show_cmp:
            st.caption("⬆️ チェックを入れると年度間比較グラフ・表を表示します。")
        if show_cmp:
            import re as _re
            _file_keys = tuple(
                (_f['name'], _f.get('modifiedTime', ''))
                for _f in drive_files
                if _re.search(r'202[0-9]', _f['name']) and _f['name'].endswith('.csv.gz')
            )
            df_monthly = load_comparison_monthly(_file_keys)

            if df_monthly.empty:
                st.warning("比較用の複数年度データが見つかりませんでした。")
            else:
                available_fy = sorted(df_monthly['fiscal_year'].unique())
                fy_colors = {str(y): c for y, c in zip(available_fy, ['#00FFFF', '#FF00FF', '#39FF14', '#FFFF00'])}

                # --- 駐車場・指標セレクタ ---
                _cmp_col1, _cmp_col2 = st.columns([2, 2])
                with _cmp_col1:
                    _area_opts = ["全駐車場"] + [p for p in PARKING_ORDER if p in df_monthly['ParkingAreaName'].unique()]
                    cmp_area = st.selectbox("駐車場を選択", _area_opts, key="cmp_area")
                with _cmp_col2:
                    cmp_metric = st.selectbox("比較指標（グラフ用）", ["利用台数", "現金収入"], key="cmp_metric")

                # フィルター
                if cmp_area != "全駐車場":
                    df_cmp = df_monthly[df_monthly['ParkingAreaName'] == cmp_area].copy()
                else:
                    df_cmp = df_monthly.copy()

                # --- KPIパネル（年度合計） ---
                st.markdown("#### 年度合計 KPI")
                kpi_cols = st.columns(len(available_fy))
                kpi_by_year = {}
                for _fy in available_fy:
                    _dfy = df_cmp[df_cmp['fiscal_year'] == _fy]
                    kpi_by_year[_fy] = {
                        '台数': int(_dfy['利用台数'].sum()),
                        '現金収入': int(_dfy['現金収入'].sum())
                    }

                for _i, _fy in enumerate(available_fy):
                    _prev_fy = available_fy[_i - 1] if _i > 0 else None
                    _cnt = kpi_by_year[_fy]['台数']
                    _cash = kpi_by_year[_fy]['現金収入']
                    _fy_color = fy_colors.get(str(_fy), '#00FFFF')
                    with kpi_cols[_i]:
                        if _prev_fy:
                            _prev_cnt = kpi_by_year[_prev_fy]['台数']
                            _prev_cash = kpi_by_year[_prev_fy]['現金収入']
                            _diff_cnt = ((_cnt - _prev_cnt) / _prev_cnt * 100) if _prev_cnt > 0 else 0
                            _diff_cash = ((_cash - _prev_cash) / _prev_cash * 100) if _prev_cash > 0 else 0
                            _ac = "▲" if _diff_cnt >= 0 else "▼"
                            _ar = "▲" if _diff_cash >= 0 else "▼"
                            _yoy = (
                                f'<div style="color:{_fy_color};font-size:16px;font-weight:700;margin-top:8px;">'
                                f'{_ac} 台数 {abs(_diff_cnt):.1f}%</div>'
                                f'<div style="color:{_fy_color};font-size:16px;font-weight:700;">'
                                f'{_ar} 収入 {abs(_diff_cash):.1f}%</div>'
                                f'<div style="color:#888;font-size:13px;">（前年比）</div>'
                            )
                        else:
                            _yoy = '<div style="color:#888;font-size:16px;margin-top:8px;">— 基準年</div>'
                        st.markdown(f"""
<div style="background:rgba(0,255,255,0.07);border:1px solid {_fy_color};border-radius:10px;padding:16px;text-align:center;">
  <div style="color:{_fy_color};font-size:18px;font-weight:700;">FY{_fy}</div>
  <div style="color:#FFFFFF;font-size:28px;font-weight:800;text-shadow:0 0 8px {_fy_color};">{_cnt:,} 台</div>
  <div style="color:#FFFFFF;font-size:28px;font-weight:800;text-shadow:0 0 8px {_fy_color};">¥{_cash:,} 円</div>
  <div style="color:#AAAAAA;font-size:12px;">（消費税含む）</div>
  {_yoy}
</div>""", unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)

                # --- 年度比較折れ線グラフ ---
                if 'Month' in df_cmp.columns:
                    _val_col = '利用台数' if cmp_metric == "利用台数" else '現金収入'
                    df_cmp_grp = df_cmp.groupby(['fiscal_year', 'Month'])[_val_col].sum().reset_index(name='値')
                    df_cmp_grp['fiscal_year'] = df_cmp_grp['fiscal_year'].astype(str)
                    _fy_month_order = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
                    _month_label_map = {m: f"{m}月" for m in range(1, 13)}
                    df_cmp_grp['月番号'] = df_cmp_grp['Month'].str[-2:].astype(int)
                    df_cmp_grp['月ラベル'] = df_cmp_grp['月番号'].map(_month_label_map)
                    df_cmp_grp['月順'] = df_cmp_grp['月番号'].map({m: i for i, m in enumerate(_fy_month_order)})
                    _x_labels = [f"{m}月" for m in _fy_month_order]
                    fig_cmp = go.Figure()
                    for _fy_s in sorted(df_cmp_grp['fiscal_year'].unique()):
                        _d = df_cmp_grp[df_cmp_grp['fiscal_year'] == _fy_s].sort_values('月順')
                        fig_cmp.add_trace(go.Scatter(
                            x=_d['月ラベル'], y=_d['値'],
                            name=f"FY{_fy_s}",
                            mode='lines+markers',
                            line=dict(color=fy_colors.get(_fy_s, '#FFFFFF'), width=3),
                            marker=dict(size=8),
                            hovertemplate=f"FY{_fy_s} %{{x}}<br>{cmp_metric}: %{{y:,}}<extra></extra>"
                        ))
                    _y_label = "利用台数（台）" if cmp_metric == "利用台数" else "現金収入（円）"
                    fig_cmp.update_layout(
                        title=dict(text=f"月別 {cmp_metric} 年度比較 — {cmp_area}", font=dict(size=18, color="#00FFFF")),
                        xaxis=dict(title="月", showgrid=True, gridcolor='rgba(255,255,255,0.1)', type='category', categoryorder='array', categoryarray=_x_labels),
                        yaxis=dict(title=_y_label, showgrid=True, gridcolor='rgba(255,255,255,0.1)', rangemode='normal'),
                        font=dict(family="sans-serif", color="#E0E0E0"),
                        plot_bgcolor="rgba(17,17,17,1)",
                        paper_bgcolor="rgba(17,17,17,1)",
                        height=550,
                        margin={'l': 30, 'r': 30, 't': 50, 'b': 30},
                        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(color="#E0E0E0"))
                    )
                    st.plotly_chart(fig_cmp, use_container_width=True)
                else:
                    st.info("Month 列が見つからないため、グラフを表示できません。")

                st.markdown("<br>", unsafe_allow_html=True)

                # --- 駐車場別年度比較表 ---
                st.markdown("#### 駐車場別 年度比較表")
                st.markdown('<p style="color:#AAAAAA;font-size:13px;margin-top:-8px;">※ 台数：千台　／　収入：百万円（消費税含む）</p>', unsafe_allow_html=True)
                _tbl_grp = df_monthly.groupby(['ParkingAreaName', 'fiscal_year']).agg(
                    利用台数=('利用台数', 'sum'), 現金収入=('現金収入', 'sum')
                ).reset_index()
                _total_grp = df_monthly.groupby('fiscal_year').agg(
                    利用台数=('利用台数', 'sum'), 現金収入=('現金収入', 'sum')
                ).reset_index()
                _total_grp['ParkingAreaName'] = '全駐車場'
                _tbl_grp = pd.concat([_tbl_grp, _total_grp], ignore_index=True)
                _pivot_cnt = _tbl_grp.pivot(index='ParkingAreaName', columns='fiscal_year', values='利用台数')
                _pivot_cash = _tbl_grp.pivot(index='ParkingAreaName', columns='fiscal_year', values='現金収入')
                _parking_row_order = [p for p in PARKING_ORDER if p in _pivot_cnt.index] + ['全駐車場']
                _fys = sorted(available_fy)

                def _yoy_cell(curr, prev, color):
                    if prev is None or prev == 0:
                        return "—"
                    pct = (curr - prev) / prev * 100
                    arrow = "▲" if pct >= 0 else "▼"
                    clr = "#39FF14" if pct >= 0 else "#FF4500"
                    return f'<span style="color:{clr};font-weight:700;">{arrow}{abs(pct):.1f}%</span>'

                _th = '<th style="background:#1a2a3a;color:#00FFFF;padding:6px 10px;text-align:center;white-space:nowrap;">'
                _th_left = '<th style="background:#1a2a3a;color:#00FFFF;padding:6px 10px;text-align:left;white-space:nowrap;">'
                header = f'<tr>{_th_left}駐車場</th>'
                for _fy in _fys:
                    header += f'{_th}FY{_fy}<br>台数</th>'
                    if _fys.index(_fy) > 0:
                        header += f'{_th}前年比</th>'
                for _fy in _fys:
                    header += f'{_th}FY{_fy}<br>収入</th>'
                    if _fys.index(_fy) > 0:
                        header += f'{_th}前年比</th>'
                header += '</tr>'
                rows_html = ""
                for _park in _parking_row_order:
                    if _park not in _pivot_cnt.index:
                        continue
                    _is_total = _park == '全駐車場'
                    _row_style = 'background:#0d1f2d;font-weight:700;border-top:2px solid #00FFFF;' if _is_total else 'background:#111827;'
                    _td = '<td style="padding:6px 10px;text-align:right;color:#E0E0E0;white-space:nowrap;">'
                    _td_left = '<td style="padding:6px 10px;text-align:left;color:#00FFFF;font-weight:700;white-space:nowrap;">'
                    row = f'<tr style="{_row_style}">{_td_left}{_park}</td>'
                    for _i, _fy in enumerate(_fys):
                        _val = _pivot_cnt.at[_park, _fy] if _fy in _pivot_cnt.columns else 0
                        row += f'{_td}{_val/1000:.1f}</td>'
                        if _i > 0:
                            _prev_val = _pivot_cnt.at[_park, _fys[_i-1]] if _fys[_i-1] in _pivot_cnt.columns else 0
                            row += f'<td style="padding:6px 10px;text-align:center;">{_yoy_cell(_val, _prev_val, fy_colors.get(str(_fy)))}</td>'
                    for _i, _fy in enumerate(_fys):
                        _val = _pivot_cash.at[_park, _fy] if _fy in _pivot_cash.columns else 0
                        row += f'{_td}{_val/1_000_000:.1f}</td>'
                        if _i > 0:
                            _prev_val = _pivot_cash.at[_park, _fys[_i-1]] if _fys[_i-1] in _pivot_cash.columns else 0
                            row += f'<td style="padding:6px 10px;text-align:center;">{_yoy_cell(_val, _prev_val, fy_colors.get(str(_fy)))}</td>'
                    row += '</tr>'
                    rows_html += row
                _table_html = f"""
<div style="overflow-x:auto;">
<table style="border-collapse:collapse;width:100%;font-size:14px;font-family:sans-serif;">
<thead>{header}</thead>
<tbody>{rows_html}</tbody>
</table>
</div>"""
                st.markdown(_table_html, unsafe_allow_html=True)

# ==========================================
# ② 24時間稼働状況分析 (原本ロジック完全再現)
# ==========================================
else:
    common_layout = dict(
        font=dict(family="sans-serif", color="#E0E0E0"),
        plot_bgcolor="rgba(17, 17, 17, 1)",
        paper_bgcolor="rgba(17, 17, 17, 1)",
        margin={'l': 30, 'r': 30, 't': 50, 'b': 30}
    )
    PARKING_CAPACITY = {"南1駐車場": 918, "南2駐車場": 601, "南3駐車場": 690, "南4駐車場": 638, "北1駐車場": 622, "北2駐車場": 248, "北3駐車場": 192, "全駐車場": 3809}
    NEON_COLORS = {"一般在庫": "#22D3EE", "定期在庫": "#F0ABFC", "在庫合計": "#4ADE80", "一般入庫": "#60A5FA", "定期入庫": "#C084FC", "一般出庫": "#F87171", "定期出庫": "#FACC15", "収容台数": "#EF4444"}

    @st.cache_data(show_spinner=True)
    def load_data_dashboard2_drive():
        """Google Driveにある '_with_avg.xlsx' ファイルを年度別に収集"""
        xls_files = search_files_in_drive("mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and name contains 'with_avg' and trashed = false")
        all_data = []
        for f_info in xls_files:
            file_name = f_info['name']
            # 年度をファイル名から抽出 (原本のフォルダ構成 [2023, 2024, 2025] に対応)
            year = 0
            for y in [2023, 2024, 2025]:
                if str(y) in file_name: year = y; break
            if year == 0: continue
            
            parts = file_name.replace("_with_avg.xlsx", "").split("_", 1)
            if len(parts) != 2: continue
            _, parking_name = parts
            if "南4B" in parking_name: continue
            
            fh = download_file_from_drive(f_info['id'])
            if not fh: continue
            try:
                for sheet in ["月平均", "平日平均", "休日平均"]:
                    df_raw = pd.read_excel(fh, sheet_name=sheet, header=None)
                    df_sub = df_raw.iloc[8:32, 3:10].copy()
                    df_sub.columns = ["時間帯", "一般入庫", "一般出庫", "一般在庫", "定期入庫", "定期出庫", "定期在庫"]
                    df_sub["年度"] = int(year); df_sub["駐車場名"] = parking_name; df_sub["曜日区分"] = sheet
                    df_sub = df_sub.dropna(subset=["時間帯"])
                    if df_sub.empty: continue
                    for col in ["一般入庫", "一般出庫", "一般在庫", "定期入庫", "定期出庫", "定期在庫"]:
                        df_sub[col] = pd.to_numeric(df_sub[col], errors='coerce').fillna(0).astype(int)
                    df_sub["在庫合計"] = df_sub["一般在庫"] + df_sub["定期在庫"]
                    all_data.append(df_sub)
            except Exception: pass
        if not all_data: return pd.DataFrame()
        return pd.concat(all_data, ignore_index=True)

    def calculate_kpis(df_selection, parking_name):
        if df_selection.empty: return None
        capacity = PARKING_CAPACITY.get(parking_name, 1); max_stock = df_selection["在庫合計"].max()
        max_occ_rate = (max_stock / capacity) * 100
        peak_time = df_selection.loc[df_selection["在庫合計"].idxmax()]["時間帯"]
        row_p = df_selection.loc[df_selection["在庫合計"].idxmax()]; total_p = row_p["在庫合計"]
        reg_dep_rate = (row_p["定期在庫"] / total_p * 100) if total_p > 0 else 0
        rot = (df_selection["一般入庫"].sum() + df_selection["一般出庫"].sum()) / (2 * capacity)
        return {"max_occ_rate": max_occ_rate, "peak_time": peak_time, "reg_dep_rate": reg_dep_rate, "traffic_activity": rot, "max_stock": max_stock}

    df2 = load_data_dashboard2_drive()
    if df2.empty: st.error("Google Drive上に分析対象のExcelデータが見つかりませんでした。")
    else:
        with st.sidebar:
            st.title("⚡ Settings")
            available_years = sorted(df2["年度"].unique())
            selected_year = st.selectbox("分析対象年度を選択", options=available_years, index=len(available_years)-1)
            st.caption("🚀 Professional Theme Active")

        df_year = df2[df2["年度"] == selected_year]
        st.title("🅿️ ② 24時間稼働状況分析")
        st.markdown(f"**{selected_year}年度** 稼働特性・KPIインサイト")
        
        # タブ名称を原本通りに設定
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📍 1. 駐車場間の波形比較", "📅 2. 平日・休日の特性比較", "🎫 3. 定期・一般の特性比較", "📈 4. 年度別の経年変化", "📊 5. 総合詳細ビュー"])
        METRIC_OPTIONS = ["在庫合計", "一般在庫", "定期在庫", "一般入庫", "一般出庫", "定期入庫", "定期出庫"]

        with tab1:
            st.subheader("📍 各駐車場の24時間稼働波形の比較")
            c1, c2 = st.columns([1, 4])
            with c1:
                t_met = st.selectbox("比較指標", options=METRIC_OPTIONS, index=0, key="t1_metric")
                d_type = st.selectbox("曜日区分", options=["月平均", "平日平均", "休日平均"], index=0, key="t1_day")
                plist = ["南1駐車場", "南2駐車場", "南3駐車場", "南4駐車場", "北1駐車場", "北2駐車場", "北3駐車場"]
                plist = [p for p in plist if p in df2["駐車場名"].unique()]
                is_all = st.checkbox("全ての駐車場を選択", value=True)
                sel_p = st.multiselect("対象駐車場", options=plist, default=plist if is_all else [])
            with c2:
                df_t1 = df_year[(df_year["曜日区分"] == d_type) & (df_year["駐車場名"].isin(sel_p))].copy()
                if not df_t1.empty:
                    fig1 = px.line(df_t1, x="時間帯", y=t_met, color="駐車場名", markers=True, line_shape="spline", category_orders={"駐車場名": plist})
                    fig1.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    fig1.update_layout(**common_layout, hovermode="x unified", height=550)
                    fig1.update_yaxes(gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    st.plotly_chart(fig1, use_container_width=True)

        with tab2:
            st.subheader("📅 各駐車場の平日と休日の稼働ギャップ")
            pk_l = sorted([p for p in df2["駐車場名"].unique() if p != "全駐車場"]) + ["全駐車場"]
            c3, c4 = st.columns([1, 4])
            with c3:
                t_pk2 = st.selectbox("駐車場を選択", options=pk_l, key="t2_pk")
                m_t2 = st.selectbox("比較指標", options=METRIC_OPTIONS, key="t2_metric")
            with c4:
                d_wd = df_year[(df_year["駐車場名"] == t_pk2) & (df_year["曜日区分"] == "平日平均")]; d_ho = df_year[(df_year["駐車場名"] == t_pk2) & (df_year["曜日区分"] == "休日平均")]
                k_wd = calculate_kpis(d_wd, t_pk2); k_ho = calculate_kpis(d_ho, t_pk2)
                if k_wd and k_ho:
                    m1, m2, m3 = st.columns(3); gap = k_ho["max_stock"] / k_wd["max_stock"] if k_wd["max_stock"] > 0 else 0
                    m1.metric("平日最大稼動率", f"{k_wd['max_occ_rate']:.1f} %", help="【計算式】(平日在庫の最大値 / 収容台数) × 100 \n\n 平日において、その駐車場の収容キャパシティに対して、ピーク時にどれだけの車両が埋まっているかを示します。この値が高い（80〜90%超）場合は、平日のビジネス・通勤需要による満車リスクが高いと判断できます。")
                    m2.metric("休日最大稼動率", f"{k_ho['max_occ_rate']:.1f} %", help="【計算式】(休日在庫の最大値 / 収容台数) × 100 \n\n 休日（土日祝）において、駐車場の収容キャパシティに対して、ピーク時にどれだけの車両が埋まっているかを示します。平日よりも高い数値を示す場合、商業施設や観光需要などのお出かけ客を主体とした運用特性であることを意味します。")
                    m3.metric("平日・休日ギャップ", f"{gap:.2f}", help="【計算式】休日最大在庫実数 / 平日最大在庫実数 \n\n 休日のピークと平日のピークの比率です。1.0を超えれば『休日混雑型（商業・レジャー系）』、1.0を大きく下回れば『平日混雑型（ビジネス・都心型拠点）』と分類できます。運営方針や割引施策の対象日を検討する重要な切り分けとなります。")
                df_t2 = pd.concat([d_wd, d_ho])
                if not df_t2.empty:
                    fig2 = px.line(df_t2, x="時間帯", y=m_t2, color="曜日区分", markers=True, color_discrete_map={"平日平均": "#00FFFF", "休日平均": "#FF00FF"})
                    if "在庫" in m_t2:
                        cap = PARKING_CAPACITY.get(t_pk2, 0)
                        if cap > 0: fig2.add_hline(y=cap, line_dash="dash", line_color=NEON_COLORS["収容台数"], annotation_text="CAPACITY")
                    fig2.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    fig2.update_layout(**common_layout, hovermode="x unified", height=500)
                    fig2.update_yaxes(gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            st.subheader("🎫 顧客構造（定期利用 vs 一般利用）の比較")
            c5, c6 = st.columns([1, 4])
            with c5:
                t_pk3 = st.selectbox("駐車場を選択", options=pk_l, key="t3_pk")
                day_t3 = st.selectbox("曜日区分選択", options=["平日平均", "休日平均", "月平均"], key="t3_day"); mod_t3 = st.radio("表示する軸", ["在庫台数", "入庫台数", "出庫台数"])
            with c6:
                df_t3 = df_year[(df_year["駐車場名"] == t_pk3) & (df_year["曜日区分"] == day_t3)]
                k_t3 = calculate_kpis(df_t3, t_pk3)
                if k_t3:
                    k1, k2, k3 = st.columns(3)
                    k1.metric("全体ピーク時刻", k_t3["peak_time"], help="【算出方法】年間平均データのうち、一般+定期の在庫合計が最大となった時刻を表示します。これがその駐車場の『最も注意が必要な時間』となります。")
                    k2.metric("定期利用依存度", f"{k_t3['reg_dep_rate']:.1f} %", help="【計算式】(ピーク時の定期在庫数 / ピーク時の在庫合計) × 100 \n\n 駐車場が満車に近づく瞬間、その利用者のうち何％が定期券利用者であるかを示します。この数値が高いほど、安定的な月極収入はあるものの、一般利用客を逃している可能性があるため、発行枚数やエリアの調整を検討する材料となります。")
                    k3.metric("一般回転率", f"{k_t3['traffic_activity']:.2f}", help="【計算式】((一般入庫合計 + 一般出庫合計) ÷ 2) ÷ 収容台数\n\n1台分のスペースが1日に平均何回入れ替わったか（回転率）を示します。数値が「1.0」であれば、1車室につき1日1台の一般客が入れ替わったことを意味します。在庫が少なくてもこの値が高ければ、短時間利用が多く高収益な拠点と言えます。")
                if not df_t3.empty:
                    fig3 = go.Figure()
                    fig3.add_trace(go.Scatter(x=df_t3["時間帯"], y=df_t3["定期在庫"] if "在庫" in mod_t3 else (df_t3["定期入庫"] if "入庫" in mod_t3 else df_t3["定期出庫"]), name="定期", mode='lines+markers', line=dict(color=NEON_COLORS["定期在庫"], width=3)))
                    fig3.add_trace(go.Scatter(x=df_t3["時間帯"], y=df_t3["一般在庫"] if "在庫" in mod_t3 else (df_t3["一般入庫"] if "入庫" in mod_t3 else df_t3["一般出庫"]), name="一般", mode='lines+markers', line=dict(color=NEON_COLORS["一般在庫"], width=3)))
                    fig3.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    fig3.update_layout(**common_layout, hovermode="x unified", height=500)
                    fig3.update_yaxes(gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    st.plotly_chart(fig3, use_container_width=True)

        with tab4:
            st.subheader("📈 年度別の稼働トレンド推移")
            c7, c8 = st.columns([1, 4])
            with c7:
                t_pk4 = st.selectbox("分析対象の駐車場", options=pk_l, key="t4_pk")
                m_t4 = st.selectbox("分析指標", options=METRIC_OPTIONS, key="t4_metric")
                d_t4_sel = st.selectbox("分析曜日", options=["月平均", "平日平均", "休日平均"], key="t4_day")
            with c8:
                df_t4 = df2[(df2["駐車場名"] == t_pk4) & (df2["曜日区分"] == d_t4_sel)]
                if not df_t4.empty:
                    fig4 = px.line(df_t4, x="時間帯", y=m_t4, color="年度", markers=True, line_shape="spline")
                    fig4.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    fig4.update_layout(**common_layout, hovermode="x unified", height=500)
                    fig4.update_yaxes(gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    st.plotly_chart(fig4, use_container_width=True)

        with tab5:
            st.subheader("📊 各駐車場総合稼働状況分析")
            c9, c10 = st.columns([1, 5])
            with c9:
                t_pk5 = st.selectbox("対象を表示", options=pk_l, key="t5_pk"); d_t5 = st.selectbox("表示タイプ", options=["月平均", "平日平均", "休日平均"], key="t5_day")
            with c10:
                df_t5_data = df_year[(df_year["駐車場名"] == t_pk5) & (df_year["曜日区分"] == d_t5)]
                k_t5 = calculate_kpis(df_t5_data, t_pk5)
                if k_t5:
                    d1, d2, d3, d4 = st.columns(4)
                    d1.metric("最大稼動ポテンシャル", f"{k_t5['max_occ_rate']:.1f} %", help="【計算式】(1日の在庫合計の最大値 ÷ 収容台数) × 100\n\n選択した駐車場・年度・曜日区分において、1日のうち最も車が多かった瞬間に、収容キャパシティの何％が埋まっていたかを示します。\n\n▶ 90%以上：満車リスクが高く、利用者が入れない機会損失が発生している可能性があります。定期券の発行枚数見直しや入庫制限の運用が検討されます。\n▶ 70〜90%：概ね適正な稼働状態です。ピーク時間帯の誘導対応が有効です。\n▶ 70%未満：余裕はあるものの、料金施策やPRによる集客強化の余地があります。")
                    d2.metric("総合ピーク在庫実数", f"{k_t5['max_stock']:,} 台", help="【算出方法】1日の全時間帯のうち「一般在庫＋定期在庫」が最大となった時刻の実際の駐車台数です。\n\nこの数値は「最大稼動ポテンシャル（%）」の分子にあたる実数であり、グラフの最高点の台数を直接確認できます。\n\n▶ 収容台数と比較することで、あと何台余裕があったかを把握できます。\n▶ 年度を切り替えて比較すると、利用台数の経年トレンドを読み取ることができます。")
                    d3.metric("定期券利用シェア", f"{k_t5['reg_dep_rate']:.1f} %", help="【計算式】(ピーク時の定期在庫台数 ÷ ピーク時の在庫合計台数) × 100\n\n駐車場が最も混雑する瞬間に、駐車している車のうち何％が定期券（月極）利用者かを示します。\n\n▶ 70%以上：月極収入は安定していますが、一般利用者を取り込む余地が少ない状態です。定期券の発行枚数を調整することで、一般収益を増やせる可能性があります。\n▶ 40〜70%：定期と一般がバランスよく共存しており、理想的な構成です。\n▶ 40%未満：一般利用主体の拠点です。利用者の入れ替わりが激しく、日によって大きく変動するリスクがあります。")
                    d4.metric("一般回転率", f"{k_t5['traffic_activity']:.2f}", help="【計算式】((1日の一般入庫合計 ＋ 1日の一般出庫合計) ÷ 2) ÷ 収容台数\n\n1日あたりに1つの駐車スペースが何回入れ替わったか（回転率）を示す指標です。\n\n▶ 1.0以上：1日で全てのスペースが1回以上入れ替わっている高回転な状態です。時間貸し収益効率が高い拠点と判断できます。\n▶ 0.5〜1.0：標準的な回転率です。\n▶ 0.5未満：入れ替わりが少なく、長時間駐車が主体、または利用者の動きが少ない拠点です。")
                if not df_t5_data.empty:
                    fig5 = make_subplots(specs=[[{"secondary_y": True}]])
                    fig5.add_trace(go.Bar(x=df_t5_data["時間帯"], y=df_t5_data["定期在庫"], name="定期在庫", marker_color="rgba(240, 171, 252, 0.5)"), secondary_y=False)
                    fig5.add_trace(go.Bar(x=df_t5_data["時間帯"], y=df_t5_data["一般在庫"], name="一般在庫", marker_color="rgba(34, 211, 238, 0.5)"), secondary_y=False)
                    fig5.add_trace(go.Scatter(x=df_t5_data["時間帯"], y=df_t5_data["一般入庫"], name="一般入庫", mode='lines+markers', line=dict(color=NEON_COLORS["一般入庫"], width=2)), secondary_y=True)
                    fig5.add_trace(go.Scatter(x=df_t5_data["時間帯"], y=df_t5_data["一般出庫"], name="一般出庫", mode='lines+markers', line=dict(color=NEON_COLORS["一般出庫"], width=2)), secondary_y=True)
                    fig5.add_trace(go.Scatter(x=df_t5_data["時間帯"], y=df_t5_data["定期入庫"], name="定期入庫", mode='lines+markers', line=dict(color=NEON_COLORS["定期入庫"], width=2)), secondary_y=True)
                    fig5.add_trace(go.Scatter(x=df_t5_data["時間帯"], y=df_t5_data["定期出庫"], name="定期出庫", mode='lines+markers', line=dict(color=NEON_COLORS["定期出庫"], width=2)), secondary_y=True)
                    cap = PARKING_CAPACITY.get(t_pk5, 0)
                    if cap > 0: fig5.add_hline(y=cap, line_dash="dash", line_color=NEON_COLORS["収容台数"], annotation_text=f"CAP ({cap})", annotation_position="left")
                    fig5.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    fig5.update_layout(**common_layout, barmode='stack', hovermode="x unified", height=650, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                    fig5.update_yaxes(title_text="在庫台数（台）", secondary_y=False, showgrid=True, gridcolor='rgba(255,255,255,0.1)', tickfont=dict(color="#E0E0E0", size=12))
                    fig5.update_yaxes(title_text="入出庫数（台/時）", secondary_y=True, showgrid=False, tickfont=dict(color="#E0E0E0", size=12))
                    st.plotly_chart(fig5, use_container_width=True)

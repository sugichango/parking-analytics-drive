import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import glob

# 0. 共通：初期設定と認証機能
# ==========================================
# スクリプトの場所を基準に絶対パスを生成（クラウド・ローカル両対応）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
st.set_page_config(
    page_title="駐車場 統合アナリティクス",
    page_icon="🅿️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- 全体デザイン＆CSS設定 ---
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
        font-weight: 700 !important;
        font-size: 16px !important;
        opacity: 1.0 !important;
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
    /* ラベル系の色修正 */
    label[data-testid="stWidgetLabel"], .stSelectbox label, .stRadio label, .stMultiSelect label {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        opacity: 1.0 !important;
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

with st.sidebar:
    st.markdown("---")
    st.subheader("🛠️ ツール切替")
    mode = st.radio("表示メニュー", ["① 一般利用台数推移分析", "② 24時間稼働状況分析"], key="sys_mode")
    if st.button("ログアウト", key="logout_act"):
        st.session_state["authenticated"] = False
        st.rerun()
    st.markdown("---")

# ==========================================
# ① 経営ダッシュボード (dashboard.py から完コピ)
# ==========================================
if "①" in mode:
    st.title("📊 ① 一般利用台数推移分析")

    @st.cache_data(show_spinner=True)
    def load_data_dashboard1(file_path):
        """CSVデータを読み込む関数（キャッシュ化とメモリ削減で高速化）"""
        use_cols = ['ParkingArea', 'OnTime', 'Cash',
                    'Discount1', 'Discount2', 'Discount3', 'Discount4', 
                    'Discount5', 'Discount6', 'Discount7']
        
        dtypes = {
            'ParkingArea': 'Int16',
            'Cash': 'Int32',
            'Discount1': 'Int16', 'Discount2': 'Int16', 'Discount3': 'Int16', 'Discount4': 'Int16',
            'Discount5': 'Int16', 'Discount6': 'Int16', 'Discount7': 'Int16'
        }
        
        try:
            header_df = pd.read_csv(file_path, nrows=0, encoding='utf-8')
            actual_cols = [c for c in use_cols if c in header_df.columns]
            actual_dtypes = {k: v for k, v in dtypes.items() if k in actual_cols}
            parse_dates = ['OnTime'] if 'OnTime' in actual_cols else False
            df = pd.read_csv(file_path, encoding='utf-8', usecols=actual_cols, dtype=actual_dtypes, parse_dates=parse_dates)
        except UnicodeDecodeError:
            try:
                header_df = pd.read_csv(file_path, nrows=0, encoding='cp932')
                actual_cols = [c for c in use_cols if c in header_df.columns]
                actual_dtypes = {k: v for k, v in dtypes.items() if k in actual_cols}
                parse_dates = ['OnTime'] if 'OnTime' in actual_cols else False
                df = pd.read_csv(file_path, encoding='cp932', usecols=actual_cols, dtype=actual_dtypes, parse_dates=parse_dates)
            except Exception as e:
                st.error(f"ファイルのエンコーディングエラー: {e}")
                return None
        except Exception as e:
            st.error(f"ファイル読み込みエラー: {e}")
            return None

        if df is not None:
            parking_areas = {440: "南１", 441: "南２", 442: "南３", 443: "南４", 444: "北１", 445: "北２", 446: "北３"}
            if 'ParkingArea' in df.columns:
                df['ParkingAreaName'] = df['ParkingArea'].map(parking_areas).fillna('不明')
                
            discount_cols = [f'Discount{i}' for i in range(1, 8)]
            active_disc_cols = [c for c in discount_cols if c in df.columns]
            for c in active_disc_cols:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype('Int64')
                
            mask_rb = pd.Series(False, index=df.index)
            mask_ticket = pd.Series(False, index=df.index)
            mask_any = pd.Series(False, index=df.index)
            
            rb_codes = [11, 12, 13, 14, 15, 43, 44]
            ticket_codes = [30, 31, 32, 33, 34, 35]
            
            for c in active_disc_cols:
                mask_rb = mask_rb | df[c].isin(rb_codes)
                mask_ticket = mask_ticket | df[c].isin(ticket_codes)
                mask_any = mask_any | (df[c] > 0)
                
            df['PaymentType'] = 'その他'
            df.loc[~mask_any, 'PaymentType'] = '現金'
            df.loc[mask_ticket & ~mask_rb, 'PaymentType'] = '回数券'
            df.loc[mask_rb, 'PaymentType'] = 'RB'
            
            if 'OnTime' in df.columns:
                df['OnTime'] = pd.to_datetime(df['OnTime'], errors='coerce')
                df['Month'] = df['OnTime'].dt.to_period('M').astype(str)
                df['DayOfWeek'] = df['OnTime'].dt.dayofweek
                df['is_holiday'] = df['DayOfWeek'].isin([5, 6]).map({True: '休日', False: '平日'})
                
        return df

    # --- スクリプトの場所を基準にした相対パス ---
    file_path = os.path.join(BASE_DIR, "data", "updated_integrated_data_FY2025.csv.gz")
    
    if os.path.exists(file_path):
        df_d1 = load_data_dashboard1(file_path)
        if df_d1 is not None:
            st.sidebar.header("🔍 フィルター設定")
            available_areas = ["全駐車場"]
            if 'ParkingAreaName' in df_d1.columns:
                areas = sorted([a for a in df_d1['ParkingAreaName'].unique() if a != '不明'])
                available_areas.extend(areas)
            selected_area = st.sidebar.selectbox("駐車場名", available_areas, index=0)

            selected_day_type = st.sidebar.selectbox("平日/休日", ["すべて", "平日", "休日"], index=0)

            available_months = ["通年"]
            if 'Month' in df_d1.columns:
                months = sorted(df_d1[df_d1['Month'] != 'NaT']['Month'].unique().tolist())
                available_months.extend(months)
            selected_month = st.sidebar.selectbox("対象月", available_months, index=0)

            filtered_df = df_d1.copy()

            if selected_area != "全駐車場" and 'ParkingAreaName' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['ParkingAreaName'] == selected_area]

            if selected_day_type != "すべて" and 'is_holiday' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['is_holiday'] == selected_day_type]

            if selected_month != "通年" and 'Month' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Month'] == selected_month]

            st.markdown(f"**現在の絞り込み**: 駐車場=`{selected_area}` | 曜日=`{selected_day_type}` | 月=`{selected_month}` (対象データ: {len(filtered_df):,} 件)")
            
            st.markdown(
                '''<div style="border: 2px solid #00FFFF; border-radius: 8px; padding: 12px; margin-top: 20px; margin-bottom: 5px; background: rgba(0, 255, 255, 0.05);">
                <b style="color: #00FFFF; font-size: 16px;">🔲 内訳表示オプション</b>
                </div>''', unsafe_allow_html=True
            )
            show_by_payment_type = st.checkbox("👉 支払い種別（現金・RB・回数券）で内訳を表示する", value=False)
            
            if not filtered_df.empty and 'OnTime' in filtered_df.columns:
                st.subheader("📈 利用台数推移 (月別または日別)")
                if 'Cash' in filtered_df.columns:
                    filtered_df['Cash'] = pd.to_numeric(filtered_df['Cash'], errors='coerce').fillna(0)
                else:
                    filtered_df['Cash'] = 0

                if selected_month == "通年":
                    x_col = 'Month'
                    x_title = '年月'
                else:
                    filtered_df['Date'] = filtered_df['OnTime'].dt.date.astype(str)
                    x_col = 'Date'
                    x_title = '日付'
                    
                fig_bar = make_subplots(specs=[[{"secondary_y": True}]])
                
                if show_by_payment_type and 'PaymentType' in filtered_df.columns:
                    bar_counts = filtered_df.groupby([x_col, 'PaymentType']).size().reset_index(name='利用台数')
                    mapping = {
                        '現金': '現金（現金のみ）',
                        '回数券': '回数券（回数券のみ、回数券+現金）',
                        'RB': 'RB（RBのみ、RB+回数券、RB+現金、RB+回数券+現金）',
                        'その他': 'その他'
                    }
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
                else:
                    bar_counts['text'] = bar_counts['利用台数'].astype(str)

                neon_colors = ['#00FFFF', '#FF00FF', '#39FF14', '#FFEA00', '#FF003C', '#9D00FF', '#00F0FF']
                common_layout = dict(
                    font=dict(family="sans-serif", color="#E0E0E0"),
                    plot_bgcolor="rgba(17, 17, 17, 1)",
                    paper_bgcolor="rgba(17, 17, 17, 1)",
                    margin={'l': 30, 'r': 30, 't': 50, 'b': 30}
                )

                fig_bar = make_subplots(specs=[[{"secondary_y": True}]])
                
                if color_col:
                    for idx, cat in enumerate(sorted(bar_counts[color_col].unique())):
                        d = bar_counts[bar_counts[color_col] == cat]
                        fig_bar.add_trace(
                            go.Bar(
                                x=d[x_col], y=d['利用台数'], name=str(cat), text=d['text'],
                                textposition='inside', insidetextanchor='middle',
                                marker_color=neon_colors[idx % len(neon_colors)]
                            ), secondary_y=False)
                    fig_bar.update_layout(barmode='stack')
                    for i, row in total_counts.iterrows():
                        fig_bar.add_annotation(
                            x=row[x_col], y=row['合計台数'], text=str(row['合計台数']),
                            showarrow=False, yshift=10, font=dict(color="#00FFFF", size=13, family="sans-serif"), yref="y"
                        )
                else:
                    fig_bar.add_trace(
                        go.Bar(x=bar_counts[x_col], y=bar_counts['利用台数'], name="利用台数", marker_color=neon_colors[0], text=bar_counts['text'], textposition='auto', textfont=dict(color="black")),
                        secondary_y=False)
                    
                line_color = '#FFFFFF'
                fig_bar.add_trace(go.Scatter(x=line_counts[x_col], y=line_counts['現金収入'], name="現金収入(全体)", mode='lines+markers', line={'color': line_color, 'width': 3}), secondary_y=True)
                fig_bar.update_layout(
                    **common_layout,
                    title=dict(text=f"{x_title} 利用台数と現金収入推移", font=dict(size=18, color="#00FFFF")),
                    xaxis_title=x_title,
                    legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, font=dict(color="#E0E0E0"))
                )
                fig_bar.update_xaxes(showgrid=True, gridcolor='rgba(255,255,255,0.1)', type='category') # X軸の西暦誤認防止！
                fig_bar.update_yaxes(title_text="利用台数（台）", secondary_y=False, rangemode='tozero', showgrid=True, gridcolor='rgba(255,255,255,0.1)')
                fig_bar.update_yaxes(title_text="現金収入（円）", secondary_y=True, rangemode='tozero', showgrid=False)
                
                st.plotly_chart(fig_bar, use_container_width=True)
                st.markdown("---")
                
                if 'ParkingAreaName' in filtered_df.columns:
                    area_counts = filtered_df.groupby('ParkingAreaName').agg({'OnTime': 'size', 'Cash': 'sum'}).rename(columns={'OnTime': '利用台数', 'Cash': '現金収入'}).reset_index()
                    total_parked = area_counts['利用台数'].sum()
                    total_cash = area_counts['現金収入'].sum()
                    parking_colors = {area: neon_colors[i % len(neon_colors)] for i, area in enumerate(sorted(filtered_df['ParkingAreaName'].unique()) if 'ParkingAreaName' in filtered_df.columns else [])}

                    if show_by_payment_type:
                        st.subheader("🍩 駐車場別 利用内訳（サンバースト図）")
                        st.write("※ グラフの要素をクリックすると、その階層を拡大（ドリルダウン）できます。中央をクリックすると元に戻ります。")
                        if 'PaymentType' in filtered_df.columns:
                            agg_df = filtered_df.groupby(['ParkingAreaName', 'PaymentType']).agg({'OnTime': 'size', 'Cash': 'sum'}).rename(columns={'OnTime': '利用台数', 'Cash': '現金収入'}).reset_index()
                            df_target = agg_df[agg_df['利用台数'] > 0]
                            fig_chart = px.sunburst(df_target, path=['ParkingAreaName', 'PaymentType'], values='利用台数', color='ParkingAreaName', color_discrete_map=parking_colors)
                            fig_chart.update_traces(texttemplate='%{label}<br>%{value}台<br>%{percentRoot}', hovertemplate='%{label}<br>利用台数: %{value}台<br>割合: %{percentRoot}', insidetextorientation='radial')
                            fig_chart.update_layout(**common_layout, height=600)
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
            else:
                st.warning("選択された条件に該当するデータがありません。フィルター条件を変更してください。")
            st.subheader("📋 抽出データプレビュー (最初の100行)")
            st.dataframe(filtered_df.head(100), use_container_width=True)
    else:
        st.error(f"データファイルが見つかりません:\n`{file_path}`\n\nファイルが正しい場所に配置されているか確認してください。")

# ==========================================
# ② 稼動分析プロ (parking_analytics_dashboard.py から完コピ)
# ==========================================
else:
    CSV_BASE_DIR = os.path.join(BASE_DIR, "data")
    
    NEON_COLORS = {
        "一般在庫": "#22D3EE", "定期在庫": "#F0ABFC", "在庫合計": "#4ADE80",
        "一般入庫": "#60A5FA", "定期入庫": "#C084FC", "一般出庫": "#F87171",
        "定期出庫": "#FACC15", "収容台数": "#EF4444",
    }
    PARKING_CAPACITY = {"南1駐車場": 918, "南2駐車場": 601, "南3駐車場": 690, "南4駐車場": 638, "北1駐車場": 622, "北2駐車場": 248, "北3駐車場": 192, "全駐車場": 3809}

    @st.cache_data(show_spinner=True)
    def load_data_dashboard2():
        all_data = []
        years = [2023, 2024, 2025]
        for year in years:
            target_dir = os.path.join(CSV_BASE_DIR, f"excel{year}_with_avg")
            if not os.path.isdir(target_dir): continue
            files = glob.glob(os.path.join(target_dir, "*.xlsx"))
            for file in files:
                filename = os.path.basename(file)
                parts = filename.replace("_with_avg.xlsx", "").split("_", 1)
                if len(parts) != 2: continue
                _, parking_name = parts
                if "南4B" in parking_name or "南４B" in parking_name: continue
                try:
                    for sheet_name in ["月平均", "平日平均", "休日平均"]:
                        df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
                        df_sub = df_raw.iloc[8:32, 3:10].copy()
                        df_sub.columns = ["時間帯", "一般入庫", "一般出庫", "一般在庫", "定期入庫", "定期出庫", "定期在庫"]
                        df_sub["年度"] = int(year); df_sub["駐車場名"] = parking_name; df_sub["曜日区分"] = sheet_name
                        df_sub = df_sub.dropna(subset=["時間帯"])
                        if df_sub.empty: continue
                        for col in ["一般入庫", "一般出庫", "一般在庫", "定期入庫", "定期出庫", "定期在庫"]:
                            df_sub[col] = pd.to_numeric(df_sub[col], errors='coerce').fillna(0).astype(int)
                        # 時間の表示を整えるための行は「追加」としてではなく、Plotlyの xaxis type=category で対応させるため、ここはいじらない
                        df_sub["在庫合計"] = df_sub["一般在庫"] + df_sub["定期在庫"]
                        all_data.append(df_sub)
                except Exception: pass
        if not all_data: return pd.DataFrame()
        return pd.concat(all_data, ignore_index=True)

    def calculate_kpis(df_selection, parking_name):
        if df_selection.empty: return None
        capacity = PARKING_CAPACITY.get(parking_name, 1)
        max_stock = df_selection["在庫合計"].max()
        max_occ_rate = (max_stock / capacity) * 100
        peak_time = df_selection.loc[df_selection["在庫合計"].idxmax()]["時間帯"]
        row_at_peak = df_selection.loc[df_selection["在庫合計"].idxmax()]
        total_at_peak = row_at_peak["在庫合計"]
        reg_dep_rate = (row_at_peak["定期在庫"] / total_at_peak * 100) if total_at_peak > 0 else 0
        traffic_activity = (df_selection["一般入庫"].sum() + df_selection["一般出庫"].sum()) / (2 * capacity)
        return {"max_occ_rate": max_occ_rate, "peak_time": peak_time, "reg_dep_rate": reg_dep_rate, "traffic_activity": traffic_activity, "max_stock": max_stock}

    df2 = load_data_dashboard2()

    if df2.empty:
         st.error(f"データファイルが見つかりません:\n`{CSV_BASE_DIR}` フォルダを確認してください。")
    else:
        with st.sidebar:
            st.title("⚡ Settings")
            available_years = sorted(df2["年度"].unique())
            selected_year = st.selectbox("分析対象年度を選択", options=available_years, index=len(available_years)-1)
            st.caption("🚀 Professional Theme Active")

        df_year = df2[df2["年度"] == selected_year]
        st.title("🅿️ ② 24時間稼働状況分析")
        st.markdown(f"**{selected_year}年度** 稼働特性・KPIインサイト")

        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📍 1. 駐車場間の波形比較", "📅 2. 平日・休日の特性比較", "🎫 3. 定期・一般の特性比較", "📈 4. 年度別の経年変化", "📊 5. 総合詳細ビュー"])
        METRIC_OPTIONS = ["在庫合計", "一般在庫", "定期在庫", "一般入庫", "一般出庫", "定期入庫", "定期出庫"]

        with tab1:
            st.subheader("📍 各駐車場の24時間稼働波形の比較")
            col_c1, col_c2 = st.columns([1, 4])
            with col_c1:
                target_metric = st.selectbox("比較指標", options=METRIC_OPTIONS, index=0, key="t1_metric")
                day_type = st.selectbox("曜日区分", options=["月平均", "平日平均", "休日平均"], index=0, key="t1_day")
                plist = sorted([p for p in df2["駐車場名"].unique() if p != "全駐車場"])
                is_all = st.checkbox("全ての駐車場を選択", value=True)
                sel_parkings = st.multiselect("対象駐車場", options=plist, default=plist if is_all else [])
            with col_c2:
                df_t1 = df_year[(df_year["曜日区分"] == day_type) & (df_year["駐車場名"].isin(sel_parkings))]
                if not df_t1.empty:
                    fig1 = px.line(df_t1, x="時間帯", y=target_metric, color="駐車場名", template="plotly_dark", markers=True, line_shape="spline")
                    fig1.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.05)')
                    fig1.update_layout(hovermode="x unified", legend_title="", height=550, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig1, use_container_width=True)

        with tab2:
            st.subheader("📅 各駐車場の平日と休日の稼働ギャップ")
            pk_list = sorted([p for p in df2["駐車場名"].unique() if p != "全駐車場"]) + ["全駐車場"]
            col_c3, col_c4 = st.columns([1, 4])
            with col_c3:
                target_pk_t2 = st.selectbox("駐車場を選択", options=pk_list, index=0, key="t2_pk")
                metric_t2 = st.selectbox("比較指標", options=METRIC_OPTIONS, index=0, key="t2_metric")
            with col_c4:
                df_t2_wd = df_year[(df_year["駐車場名"] == target_pk_t2) & (df_year["曜日区分"] == "平日平均")]
                df_t2_ho = df_year[(df_year["駐車場名"] == target_pk_t2) & (df_year["曜日区分"] == "休日平均")]
                kpi_wd = calculate_kpis(df_t2_wd, target_pk_t2); kpi_ho = calculate_kpis(df_t2_ho, target_pk_t2)
                if kpi_wd and kpi_ho:
                    m1, m2, m3 = st.columns(3)
                    gap = kpi_ho["max_stock"] / kpi_wd["max_stock"] if kpi_wd["max_stock"] > 0 else 0
                    m1.metric(
                        label="平日最大稼動率", 
                        value=f"{kpi_wd['max_occ_rate']:.1f} %", 
                        help="【計算式】(平日在庫の最大値 / 収容台数) × 100 \n\n 平日において、その駐車場の収容キャパシティに対して、ピーク時にどれだけの車両が埋まっているかを示します。この値が高い（80〜90%超）場合は、平日のビジネス・通勤需要による満車リスクが高いと判断できます。"
                    )
                    m2.metric(
                        label="休日最大稼動率", 
                        value=f"{kpi_ho['max_occ_rate']:.1f} %", 
                        help="【計算式】(休日在庫の最大値 / 収容台数) × 100 \n\n 休日（土日祝）において、駐車場の収容キャパシティに対して、ピーク時にどれだけの車両が埋まっているかを示します。平日よりも高い数値を示す場合、商業施設や観光需要などのお出かけ客を主体とした運用特性であることを意味します。"
                    )
                    m3.metric(
                        label="平日・休日ギャップ", 
                        value=f"{gap:.2f}", 
                        help="【計算式】休日最大在庫実数 / 平日最大在庫実数 \n\n 休日のピークと平日のピークの比率です。1.0を超えれば『休日混雑型（商業・レジャー系）』、1.0を大きく下回れば『平日混雑型（ビジネス・都心型拠点）』と分類できます。運営方針や割引施策の対象日を検討する重要な切り分けとなります。"
                    )
                df_t2 = pd.concat([df_t2_wd, df_t2_ho])
                if not df_t2.empty:
                    fig2 = px.line(df_t2, x="時間帯", y=metric_t2, color="曜日区分", template="plotly_dark", markers=True, color_discrete_map={"平日平均": "#00FFFF", "休日平均": "#FF00FF"})
                    if "在庫" in metric_t2:
                        cap = PARKING_CAPACITY.get(target_pk_t2, 0)
                        if cap > 0: fig2.add_hline(y=cap, line_dash="dash", line_color=NEON_COLORS["収容台数"], annotation_text="CAPACITY")
                    fig2.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.05)')
                    fig2.update_layout(hovermode="x unified", height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            st.subheader("🎫 顧客構造（定期利用 vs 一般利用）の比較")
            col_c5, col_c6 = st.columns([1, 4])
            with col_c5:
                target_pk_t3 = st.selectbox("駐車場を選択", options=pk_list, index=0, key="t3_pk")
                day_t3 = st.selectbox("曜日区分選択", options=["平日平均", "休日平均", "月平均"], index=0, key="t3_day")
                mod_t3 = st.radio("表示する軸", ["在庫台数", "入庫台数", "出庫台数"])
            with col_c6:
                df_t3 = df_year[(df_year["駐車場名"] == target_pk_t3) & (df_year["曜日区分"] == day_t3)]
                kpi_t3 = calculate_kpis(df_t3, target_pk_t3)
                if kpi_t3:
                    k1, k2, k3 = st.columns(3)
                    k1.metric(
                        label="全体ピーク時刻", 
                        value=kpi_t3["peak_time"], 
                        help="【算出方法】年間平均データのうち、一般+定期の在庫合計が最大となった時刻を表示します。これがその駐車場の『最も注意が必要な時間』となります。"
                    )
                    k2.metric(
                        label="定期利用依存度", 
                        value=f"{kpi_t3['reg_dep_rate']:.1f} %", 
                        help="【計算式】(ピーク時の定期在庫数 / ピーク時の在庫合計) × 100 \n\n 駐車場が満車に近づく瞬間、その利用者のうち何％が定期券利用者であるかを示します。この数値が高いほど、安定的な月極収入はあるものの、一般利用客を逃している可能性があるため、発行枚数やエリアの調整を検討する材料となります。"
                    )
                    k3.metric(
                        label="一般回転率", 
                        value=f"{kpi_t3['traffic_activity']:.2f}", 
                        help=(
                            "【計算式】((一般入庫合計 + 一般出庫合計) ÷ 2) ÷ 収容台数\n\n"
                            "1台分のスペースが1日に平均何回入れ替わったか（回転率）を示します。"
                            "数値が「1.0」であれば、1車室につき1日1台の一般客が入れ替わったことを意味します。"
                            "在庫が少なくてもこの値が高ければ、短時間利用が多く高収益な拠点と言えます。"
                        )
                    )
                if not df_t3.empty:
                    fig3 = go.Figure()
                    fig3.add_trace(go.Scatter(x=df_t3["時間帯"], y=df_t3["定期在庫"] if "在庫" in mod_t3 else (df_t3["定期入庫"] if "入庫" in mod_t3 else df_t3["定期出庫"]), name="定期 (Magenta)", mode='lines+markers', line=dict(color=NEON_COLORS["定期在庫"], width=3)))
                    fig3.add_trace(go.Scatter(x=df_t3["時間帯"], y=df_t3["一般在庫"] if "在庫" in mod_t3 else (df_t3["一般入庫"] if "入庫" in mod_t3 else df_t3["一般出庫"]), name="一般 (Cyan)", mode='lines+markers', line=dict(color=NEON_COLORS["一般在庫"], width=3)))
                    fig3.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.05)')
                    fig3.update_layout(template="plotly_dark", hovermode="x unified", height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig3, use_container_width=True)

        with tab4:
            st.subheader("📈 年度別の稼働トレンド推移")
            col_c7, col_c8 = st.columns([1, 4])
            with col_c7:
                target_pk_t4 = st.selectbox("分析対象の駐車場", options=pk_list, index=0, key="t4_pk")
                metric_t4 = st.selectbox("分析指標", options=METRIC_OPTIONS, index=0, key="t4_metric")
                day_t4 = st.selectbox("分析曜日", options=["月平均", "平日平均", "休日平均"], index=0, key="t4_day")
            with col_c8:
                df_t4 = df2[(df2["駐車場名"] == target_pk_t4) & (df2["曜日区分"] == day_t4)]
                if not df_t4.empty:
                    fig4 = px.line(df_t4, x="時間帯", y=metric_t4, color="年度", template="plotly_dark", markers=True, line_shape="spline")
                    fig4.update_xaxes(type='category')
                    fig4.update_layout(hovermode="x unified", height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig4, use_container_width=True)

        with tab5:
            st.subheader("📊 各駐車場総合稼働状況分析")
            col_c9, col_c10 = st.columns([1, 5])
            with col_c9:
                target_pk_t5 = st.selectbox("対象を表示", options=pk_list, index=0, key="t5_pk")
                day_t5 = st.selectbox("表示タイプ", options=["月平均", "平日平均", "休日平均"], index=0, key="t5_day")
            with col_c10:
                df_t5 = df_year[(df_year["駐車場名"] == target_pk_t5) & (df_year["曜日区分"] == day_t5)]
                kpi_t5 = calculate_kpis(df_t5, target_pk_t5)
                if kpi_t5:
                    d1, d2, d3, d4 = st.columns(4)
                    d1.metric(
                        label="最大稼動ポテンシャル",
                        value=f"{kpi_t5['max_occ_rate']:.1f} %",
                        help=(
                            "【計算式】(1日の在庫合計の最大値 ÷ 収容台数) × 100\n\n"
                            "選択した駐車場・年度・曜日区分において、1日のうち最も車が多かった瞬間に、"
                            "収容キャパシティの何％が埋まっていたかを示します。\n\n"
                            "▶ 90%以上：満車リスクが高く、利用者が入れない機会損失が発生している可能性があります。"
                            "定期券の発行枚数見直しや入庫制限の運用が検討されます。\n"
                            "▶ 70〜90%：概ね適正な稼働状態です。ピーク時間帯の誘導対応が有効です。\n"
                            "▶ 70%未満：余裕はあるものの、料金施策やPRによる集客強化の余地があります。"
                        )
                    )
                    d2.metric(
                        label="総合ピーク在庫実数",
                        value=f"{kpi_t5['max_stock']:,} 台",
                        help=(
                            "【算出方法】1日の全時間帯のうち「一般在庫＋定期在庫」が最大となった時刻の実際の駐車台数です。\n\n"
                            "この数値は「最大稼動ポテンシャル（%）」の分子にあたる実数であり、"
                            "グラフの最高点の台数を直接確認できます。\n\n"
                            "▶ 収容台数と比較することで、あと何台余裕があったかを把握できます。\n"
                            "▶ 年度を切り替えて比較すると、利用台数の経年トレンドを読み取ることができます。"
                        )
                    )
                    d3.metric(
                        label="定期券利用シェア",
                        value=f"{kpi_t5['reg_dep_rate']:.1f} %",
                        help=(
                            "【計算式】(ピーク時の定期在庫台数 ÷ ピーク時の在庫合計台数) × 100\n\n"
                            "駐車場が最も混雑する瞬間に、駐車している車のうち何％が定期券（月極）利用者かを示します。\n\n"
                            "▶ 70%以上：月極収入は安定していますが、一般利用者を取り込む余地が少ない状態です。"
                            "定期券の発行枚数を調整することで、一般収益を増やせる可能性があります。\n"
                            "▶ 40〜70%：定期と一般がバランスよく共存しており、理想的な構成です。\n"
                            "▶ 40%未満：一般利用主体の拠点です。利用者の入れ替わりが激しく、"
                            "日によって大きく変動するリスクがあります。"
                        )
                    )
                    d4.metric(
                        label="一般回転率",
                        value=f"{kpi_t5['traffic_activity']:.2f}", 
                        help=(
                            "【計算式】((1日の一般入庫合計 ＋ 1日の一般出庫合計) ÷ 2) ÷ 収容台数\n\n"
                            "1日あたりに1つの駐車スペースが何回入れ替わったか（回転率）を示す指標です。\n\n"
                            "▶ 1.0以上：1日で全てのスペースが1回以上入れ替わっている高回転な状態です。"
                            "時間貸し収益効率が高い拠点と判断できます。\n"
                            "▶ 0.5〜1.0：標準的な回転率です。\n"
                            "▶ 0.5未満：入れ替わりが少なく、長時間駐車が主体、または利用者の動きが少ない拠点です。"
                        )
                    )
                if not df_t5.empty:
                    fig5 = go.Figure()
                    fig5.add_trace(go.Bar(x=df_t5["時間帯"], y=df_t5["定期在庫"], name="定期在庫", marker_color="rgba(240, 171, 252, 0.5)"))
                    fig5.add_trace(go.Bar(x=df_t5["時間帯"], y=df_t5["一般在庫"], name="一般在庫", marker_color="rgba(34, 211, 238, 0.5)"))
                    fig5.add_trace(go.Scatter(x=df_t5["時間帯"], y=df_t5["一般入庫"], name="一般入庫", mode='lines+markers', line=dict(color=NEON_COLORS["一般入庫"], width=2)))
                    fig5.add_trace(go.Scatter(x=df_t5["時間帯"], y=df_t5["一般出庫"], name="一般出庫", mode='lines+markers', line=dict(color=NEON_COLORS["一般出庫"], width=2)))
                    fig5.add_trace(go.Scatter(x=df_t5["時間帯"], y=df_t5["定期入庫"], name="定期入庫", mode='lines+markers', line=dict(color=NEON_COLORS["定期入庫"], width=2)))
                    fig5.add_trace(go.Scatter(x=df_t5["時間帯"], y=df_t5["定期出庫"], name="定期出庫", mode='lines+markers', line=dict(color=NEON_COLORS["定期出庫"], width=2)))
                    cap = PARKING_CAPACITY.get(target_pk_t5, 0)
                    if cap > 0: fig5.add_hline(y=cap, line_dash="dash", line_color=NEON_COLORS["収容台数"], annotation_text=f"CAP ({cap})")
                    fig5.update_xaxes(type='category')
                    fig5.update_layout(template="plotly_dark", barmode='stack', hovermode="x unified", height=650, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                    st.plotly_chart(fig5, use_container_width=True)

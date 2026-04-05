import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# ==========================================
# 定数・初期設定
# ==========================================
st.set_page_config(
    page_title="駐車場稼働 プロフェッショナル・アナリティクス",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 【究極のプロ仕様】最高視認性の「？」マークと全体デザインCSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
    
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #0E1117;
        color: #FFFFFF;
        font-family: 'Inter', sans-serif;
    }
    
    /* 1. KPIカード領域 */
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

    /* ヘルプアイコン (?) の抜本的な視認性修正：文字として描画 */
    [data-testid="stMetricLabel"] button {
        background-color: #FFFF00 !important; /* 鮮やかなイエロー背景 */
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
    
    /* 元のSVGアイコンを完全に消す */
    [data-testid="stMetricLabel"] button svg {
        display: none !important;
    }

    /* 直接「？」という文字を太く大きな黒色で描写 */
    [data-testid="stMetricLabel"] button::before {
        content: "?" !important;
        color: #000000 !important;
        font-size: 22px !important;
        font-weight: 900 !important;
        font-family: 'Arial Black', sans-serif !important;
        display: block !important;
        line-height: 1 !important;
    }

    /* 2. 入力ウィジェット（セレクトボックス等） */
    label[data-testid="stWidgetLabel"], 
    .stSelectbox label, 
    .stRadio label, 
    .stMultiSelect label {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        opacity: 1.0 !important;
    }
    
    /* その他全体設定 */
    h1, h2, h3 { color: #FFFFFF; font-weight: 800; }
</style>
""", unsafe_allow_html=True)

CSV_BASE_DIR = r"C:\Users\sugitamasahiko\Documents\parking_system\downloads\csv"

# 色の設定
NEON_COLORS = {
    "一般在庫": "#22D3EE", "定期在庫": "#F0ABFC", "在庫合計": "#4ADE80",
    "一般入庫": "#60A5FA", "定期入庫": "#C084FC", "一般出庫": "#F87171",
    "定期出庫": "#FACC15", "収容台数": "#EF4444",
}

PARKING_CAPACITY = {
    "南1駐車場": 918, "南2駐車場": 601, "南3駐車場": 690, "南4駐車場": 638,
    "北1駐車場": 622, "北2駐車場": 248, "北3駐車場": 192, "全駐車場": 3809,
}

# ==========================================
# データロードモジュール
# ==========================================
@st.cache_data(show_spinner=True)
def load_data():
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
    traffic_activity = (df_selection["一般入庫"].sum() + df_selection["一般出庫"].sum()) / capacity
    return {"max_occ_rate": max_occ_rate, "peak_time": peak_time, "reg_dep_rate": reg_dep_rate, "traffic_activity": traffic_activity, "max_stock": max_stock}

df = load_data()

with st.sidebar:
    st.title("⚡ Settings")
    available_years = sorted(df["年度"].unique())
    selected_year = st.selectbox("分析対象年度を選択", options=available_years, index=len(available_years)-1)
    st.markdown("---")
    st.caption("🚀 Professional Theme Active")
    st.caption("Mode: Neon Dark / #0E1117")

df_year = df[df["年度"] == selected_year]
st.title("🅿️ Parking Analytics Dashboard")
st.markdown(f"**{selected_year}年度** 稼働特性・KPIインサイト")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📍 1. 駐車場間の波形比較", "📅 2. 平日・休日の特性比較", "🎫 3. 定期・一般の特性比較", "📈 4. 年度別の経年変化", "📊 5. 総合詳細ビュー"])
METRIC_OPTIONS = ["在庫合計", "一般在庫", "定期在庫", "一般入庫", "一般出庫", "定期入庫", "定期出庫"]

with tab1:
    st.subheader("📍 各拠点の24時間稼働波形の比較")
    col_c1, col_c2 = st.columns([1, 4])
    with col_c1:
        target_metric = st.selectbox("比較指標", options=METRIC_OPTIONS, index=0, key="t1_metric")
        day_type = st.selectbox("曜日区分", options=["月平均", "平日平均", "休日平均"], index=0, key="t1_day")
        plist = sorted([p for p in df["駐車場名"].unique() if p != "全駐車場"])
        sel_parkings = st.multiselect("対象駐車場", options=plist, default=plist[:4])
    with col_c2:
        df_t1 = df_year[(df_year["曜日区分"] == day_type) & (df_year["駐車場名"].isin(sel_parkings))]
        if not df_t1.empty:
            fig1 = px.line(df_t1, x="時間帯", y=target_metric, color="駐車場名", template="plotly_dark", markers=True, line_shape="spline")
            fig1.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.05)'); fig1.update_yaxes(gridcolor='rgba(255,255,255,0.05)')
            fig1.update_layout(hovermode="x unified", legend_title="", height=550, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig1, use_container_width=True)

with tab2:
    st.subheader("📅 同一拠点における平日と休日の稼働ギャップ")
    pk_list = sorted([p for p in df["駐車場名"].unique() if p != "全駐車場"]) + ["全駐車場"]
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
            fig2.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.05)'); fig2.update_layout(hovermode="x unified", height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
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
                label="一般活性度", 
                value=f"{kpi_t3['traffic_activity']:.2f}", 
                help="【計算式】(1日の一般入庫合計台数 + 1日の一般出庫合計台数) / 駐車場の収容台数 \n\n 1日を通じた『一般利用客の入れ替わり（回転率）』のポテンシャルを示します。在庫が少なくてもこの値が高ければ、土地あたりの利用頻度が高く、高収益な回転型拠点の特性を示唆します。"
            )

        if not df_t3.empty:
            fig3 = go.Figure()
            fig3.add_trace(go.Scatter(x=df_t3["時間帯"], y=df_t3["定期在庫"] if "在庫" in mod_t3 else (df_t3["定期入庫"] if "入庫" in mod_t3 else df_t3["定期出庫"]), name="定期 (Magenta)", mode='lines+markers', line=dict(color=NEON_COLORS["定期在庫"], width=3)))
            fig3.add_trace(go.Scatter(x=df_t3["時間帯"], y=df_t3["一般在庫"] if "在庫" in mod_t3 else (df_t3["一般入庫"] if "入庫" in mod_t3 else df_t3["一般出庫"]), name="一般 (Cyan)", mode='lines+markers', line=dict(color=NEON_COLORS["一般在庫"], width=3)))
            fig3.update_xaxes(type='category', gridcolor='rgba(255,255,255,0.05)'); fig3.update_layout(template="plotly_dark", hovermode="x unified", height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig3, use_container_width=True)

with tab4:
    st.subheader("📈 年度別の稼働トレンド推移")
    col_c7, col_c8 = st.columns([1, 4])
    with col_c7:
        target_pk_t4 = st.selectbox("分析対象の駐車場", options=pk_list, index=0, key="t4_pk")
        metric_t4 = st.selectbox("分析指標", options=METRIC_OPTIONS, index=0, key="t4_metric")
        day_t4 = st.selectbox("分析曜日", options=["月平均", "平日平均", "休日平均"], index=0, key="t4_day")
    with col_c8:
        df_t4 = df[(df["駐車場名"] == target_pk_t4) & (df["曜日区分"] == day_t4)]
        if not df_t4.empty:
            fig4 = px.line(df_t4, x="時間帯", y=metric_t4, color="年度", template="plotly_dark", markers=True, color_discrete_sequence=["#00FFFF", "#FF00FF", "#00FF00"])
            fig4.update_xaxes(type='category'); fig4.update_layout(hovermode="x unified", height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig4, use_container_width=True)

with tab5:
    st.subheader("📊 拠点別・統合ライフサイクル分析")
    col_c9, col_c10 = st.columns([1, 5])
    with col_c9:
        target_pk_t5 = st.selectbox("対象を表示", options=pk_list, index=0, key="t5_pk")
        day_t5 = st.selectbox("表示タイプ", options=["月平均", "平日平均", "休日平均"], index=0, key="t5_day")
    with col_c10:
        df_t5 = df_year[(df_year["駐車場名"] == target_pk_t5) & (df_year["曜日区分"] == day_t5)]
        kpi_t5 = calculate_kpis(df_t5, target_pk_t5)
        if kpi_t5:
            d1, d2, d3, d4 = st.columns(4)
            d1.metric(label="最大稼動ポテンシャル", value=f"{kpi_t5['max_occ_rate']:.1f} %", help="【計算式】(在庫最大値/収容台数)×100")
            d2.metric(label="総合ピーク在庫実数", value=f"{kpi_t5['max_stock']:,} 台", help="1日の最大在庫数です。")
            d3.metric(label="定期券利用シェア", value=f"{kpi_t5['reg_dep_rate']:.1f} %", help="【計算式】(ピーク時定期/ピーク時合計)×100")
            d4.metric(label="一般客回転指数", value=f"{kpi_t5['traffic_activity']:.2f}", help="一般入出庫の回転率です。")
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
            fig5.update_xaxes(type='category'); fig5.update_layout(template="plotly_dark", barmode='stack', hovermode="x unified", height=650, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig5, use_container_width=True)

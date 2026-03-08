import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# ページ設定
st.set_page_config(page_title="駐車場 経営ダッシュボード(一般利用)", layout="wide")

st.title("📊 駐車場 経営ダッシュボード(一般利用)")

# --- 認証機能の追加 ---
def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        email = st.session_state["email_input"].strip()
        
        # 1. 許可されたドメイン（@tutc.or.jp）かどうかチェック
        if not email.endswith("@tutc.or.jp"):
            st.session_state["password_correct"] = False
            st.error("許可されていないメールドメインです（@tutc.or.jp を使用してください）。")
            return
            
        # 2. パスワードのチェック
        if st.session_state["password_input"] == st.secrets.get("app_password", ""):
            st.session_state["password_correct"] = True
            # 安全のため入力されたパスワードを削除
            del st.session_state["password_input"]  
        else:
            st.session_state["password_correct"] = False
            st.error("パスワードが間違っています。")

    if st.session_state.get("password_correct", False):
        return True

    # 認証用フォームの表示
    st.markdown("### 🔒 ダッシュボードへのログイン")
    st.text_input("メールアドレス（例: test@tutc.or.jp）", key="email_input")
    st.text_input("パスワード", type="password", key="password_input")
    st.button("ログイン", on_click=password_entered)
    return False

# 認証が通っていない場合はダッシュボードの表示を中止する
if not check_password():
    st.stop()

@st.cache_data
def load_data(file_path):
    """CSVデータを読み込む関数（キャッシュ化とメモリ削減で高速化）"""
    # メモリ節約のため必要なカラムだけを読み込むように指定
    use_cols = ['ParkingArea', 'OnTime', 'Cash',
                'Discount1', 'Discount2', 'Discount3', 'Discount4', 
                'Discount5', 'Discount6', 'Discount7']
    
    try:
        # まずファイルのヘッダーだけ読んで存在するカラムのみ抽出
        header_df = pd.read_csv(file_path, nrows=0, encoding='utf-8')
        actual_cols = [c for c in use_cols if c in header_df.columns]
        
        df = pd.read_csv(file_path, encoding='utf-8', usecols=actual_cols)
    except UnicodeDecodeError:
        try:
            header_df = pd.read_csv(file_path, nrows=0, encoding='cp932')
            actual_cols = [c for c in use_cols if c in header_df.columns]
            df = pd.read_csv(file_path, encoding='cp932', usecols=actual_cols)
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

# データファイルのパスを設定 (容量制限を回避するため、圧縮済みの全件データを読み込みます)
# read_csv は .gz 拡張子から自動的に解凍処理(gzip)を行います
file_path = "updated_integrated_data_FY2025.csv.gz"

if os.path.exists(file_path):
    df = load_data(file_path)
    
    if df is not None:
        
        # --- フィルターの設定 ---
        st.sidebar.header("🔍 フィルター設定")
        
        # 1. データ前処理は load_data 内へ移動済み

        # --- フィルター要素の配置 ---
        # 1. 駐車場名
        available_areas = ["全駐車場"]
        if 'ParkingAreaName' in df.columns:
            areas = sorted([a for a in df['ParkingAreaName'].unique() if a != '不明'])
            available_areas.extend(areas)
        selected_area = st.sidebar.selectbox("駐車場名", available_areas, index=0)

        # 2. 平日・休日
        selected_day_type = st.sidebar.selectbox(
            "平日/休日", 
            ["すべて", "平日", "休日"], 
            index=0
        )

        # 3. 対象月
        available_months = ["通年"]
        if 'Month' in df.columns:
            # NaTを取り除きソート
            months = sorted(df[df['Month'] != 'NaT']['Month'].unique().tolist())
            available_months.extend(months)
        selected_month = st.sidebar.selectbox("対象月", available_months, index=0)

        # --- データのフィルタリング適用 ---
        filtered_df = df.copy()

        # 駐車場フィルター
        if selected_area != "全駐車場" and 'ParkingAreaName' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['ParkingAreaName'] == selected_area]

        # 平日/休日フィルター
        if selected_day_type != "すべて" and 'is_holiday' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['is_holiday'] == selected_day_type]

        # 対象月フィルター
        if selected_month != "通年" and 'Month' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Month'] == selected_month]

        # --- フィルタリング結果の表示 ---
        st.markdown(f"**現在の絞り込み**: 駐車場=`{selected_area}` | 曜日=`{selected_day_type}` | 月=`{selected_month}` (対象データ: {len(filtered_df):,} 件)")
        
        show_by_payment_type = st.checkbox("支払い種別（現金・RB・回数券）で内訳を表示する", value=False)
        
        # グラフ用のデータが存在するか確認
        if not filtered_df.empty and 'OnTime' in filtered_df.columns:
            
            st.subheader("📈 利用台数推移 (月別または日別)")
            
            # 列にCashがあることを確認し、数値化しておく
            if 'Cash' in filtered_df.columns:
                filtered_df['Cash'] = pd.to_numeric(filtered_df['Cash'], errors='coerce').fillna(0)
            else:
                filtered_df['Cash'] = 0

            # 月が絞り込まれている場合は「日別」、通年の場合は「月別」で集計を変える工夫
            if selected_month == "通年":
                x_col = 'Month'
                x_title = '年月'
            else:
                # 対象月が選ばれている場合は日付単位(YYYY-MM-DD)で集計
                filtered_df['Date'] = filtered_df['OnTime'].dt.date.astype(str)
                x_col = 'Date'
                x_title = '日付'
                
            fig_bar = make_subplots(specs=[[{"secondary_y": True}]])
            
            if show_by_payment_type and 'PaymentType' in filtered_df.columns:
                # 支払い種別での台数集計（積み上げ表示用）
                bar_counts = filtered_df.groupby([x_col, 'PaymentType']).size().reset_index(name='利用台数')
                # 凡例用の名称マッピング
                mapping = {
                    '現金': '現金（現金のみ）',
                    '回数券': '回数券（回数券のみ、回数券+現金）',
                    'RB': 'RB（RBのみ、RB+回数券、RB+現金、RB+回数券+現金）',
                    'その他': 'その他'
                }
                bar_counts['PaymentTypeLegend'] = bar_counts['PaymentType'].map(mapping).fillna(bar_counts['PaymentType'])
                color_col = 'PaymentTypeLegend'
            elif selected_area == "全駐車場" and 'ParkingAreaName' in filtered_df.columns:
                # 駐車場別の台数集計（積み上げ表示用）
                bar_counts = filtered_df.groupby([x_col, 'ParkingAreaName']).size().reset_index(name='利用台数')
                color_col = 'ParkingAreaName'
            else:
                bar_counts = filtered_df.groupby(x_col).size().reset_index(name='利用台数')
                color_col = None
            
            # 全体の現金収入集計（折れ線グラフ用）
            line_counts = filtered_df.groupby(x_col).agg({'Cash': 'sum'}).rename(columns={'Cash': '現金収入'}).reset_index()
            
            # --- グラフ描画前のデータ加工（合計台数とパーセント計算） ---
            total_counts = bar_counts.groupby(x_col)['利用台数'].sum().reset_index(name='合計台数')
            if color_col:
                bar_counts = pd.merge(bar_counts, total_counts, on=x_col)
                bar_counts['割合'] = (bar_counts['利用台数'] / bar_counts['合計台数'] * 100).round(1)
                # 割合が0より大きい場合のみテキストを表示（見やすさ考慮）
                bar_counts['text'] = bar_counts.apply(lambda row: f"{row['割合']}%" if row['割合'] > 0 else "", axis=1)
            else:
                bar_counts['text'] = bar_counts['利用台数'].astype(str)

            # --- デザイン・テーマ設定（ダーク・ネオン調） ---
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
                            x=d[x_col], 
                            y=d['利用台数'], 
                            name=str(cat),
                            text=d['text'],
                            textposition='inside',
                            insidetextanchor='middle',
                            marker_color=neon_colors[idx % len(neon_colors)]
                        ),
                        secondary_y=False,
                    )
                fig_bar.update_layout(barmode='stack')
                
                # 棒グラフのてっぺんに合計台数を表示
                for i, row in total_counts.iterrows():
                    fig_bar.add_annotation(
                        x=row[x_col],
                        y=row['合計台数'],
                        text=str(row['合計台数']),
                        showarrow=False,
                        yshift=10,
                        font=dict(color="#00FFFF", size=13, family="sans-serif"),
                        yref="y"
                    )
            else:
                fig_bar.add_trace(
                    go.Bar(
                        x=bar_counts[x_col], 
                        y=bar_counts['利用台数'], 
                        name="利用台数", 
                        marker_color=neon_colors[0],
                        text=bar_counts['text'],
                        textposition='auto',
                        textfont=dict(color="black")
                    ),
                    secondary_y=False,
                )
                
            # 合計の現金収入を折れ線グラフで追加
            # 棒グラフの色（シアン、マゼンタ、緑等）と被りにくい、目立つ白色または明るいオレンジ色を設定
            line_color = '#FFFFFF' # 白
            fig_bar.add_trace(
                go.Scatter(x=line_counts[x_col], y=line_counts['現金収入'], name="現金収入(全体)", mode='lines+markers', line={'color': line_color, 'width': 3}),
                secondary_y=True,
            )
            
            fig_bar.update_layout(
                **common_layout,
                title=dict(text=f"{x_title} 利用台数と現金収入推移", font=dict(size=18, color="#00FFFF")),
                xaxis_title=x_title,
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.2,
                    xanchor="center",
                    x=0.5,
                    font=dict(color="#E0E0E0")
                )
            )
            
            fig_bar.update_xaxes(showgrid=True, gridcolor='rgba(255,255,255,0.1)')
            fig_bar.update_yaxes(title_text="利用台数（台）", secondary_y=False, rangemode='tozero', showgrid=True, gridcolor='rgba(255,255,255,0.1)')
            fig_bar.update_yaxes(title_text="現金収入（円）", secondary_y=True, rangemode='tozero', showgrid=False)
            
            # メイン画面に表示
            # 推移グラフ
            st.plotly_chart(fig_bar, use_container_width=True)
            
            st.markdown("---")
            
            # --- 駐車場別利用台数と現金収入のグラフ ---
            if 'ParkingAreaName' in filtered_df.columns:
                # 駐車場ごとに利用台数と現金収入を集計
                area_counts = filtered_df.groupby('ParkingAreaName').agg({'OnTime': 'size', 'Cash': 'sum'}).rename(columns={'OnTime': '利用台数', 'Cash': '現金収入'}).reset_index()
                
                # 全体の利用台数と現金収入
                total_parked = area_counts['利用台数'].sum()
                total_cash = area_counts['現金収入'].sum()
                
                # サンバースト用のネオンカラー割り当て
                # 駐車場名に色を割り当てるためのダミーマッピング（カテゴリ数に合わせてループ）
                parking_colors = {area: neon_colors[i % len(neon_colors)] for i, area in enumerate(sorted(filtered_df['ParkingAreaName'].unique()) if 'ParkingAreaName' in filtered_df.columns else [])}

                if show_by_payment_type:
                    st.subheader("🍩 駐車場別 利用内訳（サンバースト図）")
                    st.write("※ グラフの要素をクリックすると、その階層を拡大（ドリルダウン）できます。中央をクリックすると元に戻ります。")
                    
                    # 階層の準備
                    if 'PaymentType' in filtered_df.columns:
                        agg_df = filtered_df.groupby(['ParkingAreaName', 'PaymentType']).agg({'OnTime': 'size', 'Cash': 'sum'}).rename(columns={'OnTime': '利用台数', 'Cash': '現金収入'}).reset_index()
                        
                        # 指示: チェックボックスがオンの時は「現金収入」は含めない (利用台数のみを表示)
                        df_target = agg_df[agg_df['利用台数'] > 0]
                        fig_chart = px.sunburst(
                            df_target,
                            path=['ParkingAreaName', 'PaymentType'],
                            values='利用台数',
                            color='ParkingAreaName', # 駐車場ごとに色分け
                            color_discrete_map=parking_colors
                        )
                        # カスタムデータを使用して単位「台」を追加
                        fig_chart.update_traces(
                            texttemplate='%{label}<br>%{value}台<br>%{percentRoot}',
                            hovertemplate='%{label}<br>利用台数: %{value}台<br>割合: %{percentRoot}',
                            insidetextorientation='radial'
                        )
                        fig_chart.update_layout(
                            **common_layout,
                            height=600
                        )
                else:
                    st.subheader("🍩 駐車場別 利用割合 ＆ 現金収入割合")
                    
                    fig_chart = go.Figure()
                    
                    # ドーナツの表示領域を調整（現金収入が常に一番外側）
                    inner_domain = [0.15, 0.85]
                    middle_domain = [0.0, 1.0]   # 現金収入
                    inner_hole = 0.55
                    middle_hole = 0.8

                    # 内側: 利用台数
                    fig_chart.add_trace(go.Pie(
                        labels=area_counts['ParkingAreaName'],
                        values=area_counts['利用台数'],
                        name="利用台数",
                        hole=inner_hole,
                        domain={'x': inner_domain, 'y': inner_domain},
                        hoverinfo='label+value+percent+name',
                        textinfo='label+value+percent',
                        textposition='inside',
                        direction='clockwise',
                        sort=False,
                        marker=dict(colors=[parking_colors.get(label, '#FFFFFF') for label in area_counts['ParkingAreaName']]),
                        insidetextfont=dict(color="black")
                    ))
                    
                    # 一番外側: 現金収入
                    fig_chart.add_trace(go.Pie(
                        labels=area_counts['ParkingAreaName'],
                        values=area_counts['現金収入'],
                        name="現金収入",
                        hole=middle_hole,
                        domain={'x': middle_domain, 'y': middle_domain},
                        hoverinfo='label+value+percent+name',
                        textinfo='value+percent',
                        textposition='inside',
                        direction='clockwise',
                        sort=False,
                        marker=dict(colors=[parking_colors.get(label, '#FFFFFF') for label in area_counts['ParkingAreaName']]),
                        insidetextfont=dict(color="black")
                    ))
                    
                    # ドーナツの中心に「全体利用台数」と「全体現金収入」をアノテーションとして追加
                    fig_chart.update_layout(
                        **common_layout,
                        title=dict(text="内側: 利用台数 / 外側: 現金収入", font=dict(color="#00FFFF")),
                        annotations=[{
                            "text": f"総台数<br><b style='font_size:20px;'>{total_parked:,}</b><br>台<br><br>総現金<br><b style='font_size:16px;'>{int(total_cash):,}</b><br>円", 
                            "x": 0.5, 
                            "y": 0.5, 
                            "showarrow": False,
                            "font": dict(color="#00FFFF")
                        }],
                        showlegend=True,
                        legend=dict(font=dict(color="#E0E0E0"))
                    )

                st.plotly_chart(fig_chart, use_container_width=True)
                st.markdown("---")

            
        elif filtered_df.empty:
            st.warning("選択された条件に該当するデータがありません。フィルター条件を変更してください。")
        else:
            st.warning("データに 'OnTime' 列が見つからないため、グラフを描画できません。")

        # データプレビュー
        st.subheader("📋 抽出データプレビュー (最初の100行)")
        
        try:
            st.dataframe(filtered_df.head(100), use_container_width=True)
        except Exception:
            st.dataframe(filtered_df.head(100))
            
else:
    st.error(f"データファイルが見つかりません: `{file_path}`\n\nファイルが正しい場所に配置されているか確認してください。")

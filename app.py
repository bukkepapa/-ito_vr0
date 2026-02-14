import streamlit as st
import pandas as pd
from datetime import datetime
import yaml
import io
import openpyxl
from streamlit_sortables import sort_items
from utils import load_customer_data, optimize_route, calculate_schedule, get_distance_matrix, haversine

# ページ設定
st.set_page_config(page_title="自販機訪問管理表作成アプリ", layout="wide")

# CSSによるスタイル調整
st.markdown("""
<style>
    .stButton>button {
        width: 100%;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
    }
    div[data-testid="stExpander"] div[role="button"] p {
        font-size: 1.1rem;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# 設定読み込み
def load_config():
    with open('config.yaml', 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)

CONFIG = load_config()

# セッション状態の初期化
if 'master_df' not in st.session_state:
    st.session_state['master_df'] = pd.DataFrame()
if 'today_list' not in st.session_state:
    st.session_state['today_list'] = [] # list of dicts
if 'optimized_route' not in st.session_state:
    st.session_state['optimized_route'] = [] # list of dicts (customer data)

# サイドバー設定
st.sidebar.title("設定")

origin_address = st.sidebar.text_input("起点住所", value=CONFIG['defaults']['origin_address'])
destination_address = st.sidebar.text_input("終点住所 (未入力で起点と同一)", value=CONFIG['defaults']['destination_address'])
if not destination_address:
    destination_address = origin_address

departure_time_str = st.sidebar.time_input("出発時刻", value=datetime.strptime(CONFIG['defaults']['departure_time'], "%H:%M").time())
work_minutes_def = st.sidebar.number_input("標準作業時間(分)", value=CONFIG['defaults']['work_minutes'], min_value=1)

lunch_start = st.sidebar.time_input("昼休憩開始", value=datetime.strptime(CONFIG['defaults']['lunch_start'], "%H:%M").time())
lunch_end = st.sidebar.time_input("昼休憩終了", value=datetime.strptime(CONFIG['defaults']['lunch_end'], "%H:%M").time())

api_key = st.sidebar.text_input("Google Maps API Key", value=CONFIG['google_maps_api_key'])

# メインレイアウト
st.title("自販機訪問管理表作成アプリ (MVP)")

# ファイルアップロード
uploaded_file = st.file_uploader("顧客マスタをアップロード (Excel/CSV)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    # 毎回読み込むと重いので、ファイルが変わった時だけ読み込む制御を入れたいが
    # MVPではシンプルに読み込む（あるいは前回と同じならスキップ）
    # ここでは簡易実装として再読込
    df, error = load_customer_data(uploaded_file)
    if error:
        st.error(error)
    else:
        st.session_state['master_df'] = df
        st.success(f"{len(df)}件の顧客データを読み込みました。")

# 2ペイン構成
col1, col2 = st.columns([1, 1])

with col1:
    st.header("① 顧客リスト")
    if not st.session_state['master_df'].empty:
        # リスト欄が狭いという要望に対応し、データフレームを表示して視認性を高める
        st.dataframe(st.session_state['master_df'], height=300)
        
        search_query = st.text_input("検索 (コード/名称)", "")
        
        filtered_df = st.session_state['master_df'].copy()
        if search_query:
            filtered_df = filtered_df[
                filtered_df['code'].astype(str).str.contains(search_query) | 
                filtered_df['name'].str.contains(search_query)
            ]
        
        # 並び替え処理
        sort_option = st.radio("並び替え", ["コード順", "売上見込順"], horizontal=True)
        if sort_option == "売上見込順":
            filtered_df = filtered_df.sort_values(by='sales', ascending=False)
        else:
            filtered_df = filtered_df.sort_values(by='code', ascending=True)
        
        # 選択用リスト表示
        # streamlit-sortablesを使うには、リスト形式で渡す必要がある
        # ここではリストから選択して「追加」ボタンで右に移す方式（要件の代替UI）を採用
        # D&DはSortablesだと「並び替え」には強いが、「2つのリスト間の移動」は標準コンポーネントのみでは少し複雑なため
        
        # マルチセレクトで代用（検索と相性が良い）
        # 表示名を工夫: "コード : 名称 (¥売上)"
        options = filtered_df.apply(lambda x: f"{x['code']} : {x['name']} (¥{x['sales']:,})", axis=1).tolist()
        selected_items = st.multiselect("訪問先を選択", options)
        
        if st.button("TODAYリストへ追加"):
            current_codes = [item['code'] for item in st.session_state['today_list']]
            added_count = 0
            for item_str in selected_items:
                code = item_str.split(' : ')[0]
                if code not in current_codes:
                    if len(st.session_state['today_list']) >= CONFIG['defaults']['max_today_items']:
                        st.warning(f"30件の上限に達しました。")
                        break
                    
                    row = st.session_state['master_df'][st.session_state['master_df']['code'].astype(str) == code].iloc[0]
                    item_dict = row.to_dict()
                    item_dict['MUST'] = False # MUSTフラグ初期化
                    st.session_state['today_list'].append(item_dict)
                    added_count += 1
            
            if added_count > 0:
                st.success(f"{added_count}件追加しました。")
                st.rerun()

with col2:
    st.header("② TODAYリスト")
    
    if st.session_state['today_list']:
        st.write(f"Debug: first item sales = {st.session_state['today_list'][0].get('sales')} type={type(st.session_state['today_list'][0].get('sales'))}")
        # st.write(st.session_state['today_list'])
    
    total_sales = sum([int(item.get('sales', 0)) for item in st.session_state['today_list']])
    st.metric("合計売上見込", f"¥{total_sales:,}")
    
    if st.session_state['today_list']:
        # リスト編集機能（data_editor）
        df_today = pd.DataFrame(st.session_state['today_list'])
        
        # 列の並び順と表示設定
        # 必須列があるか確認
        if 'MUST' not in df_today.columns:
            df_today['MUST'] = False
            
        # 表示したい列を定義
        display_cols = ['MUST', 'code', 'name', 'sales', 'WorkMinutes', 'NoEntryTime', 'address', 'lat', 'lng']
        # 存在しない列は除外
        display_cols = [c for c in display_cols if c in df_today.columns]
        
        st.info("作業時間を編集できます")

        edited_df = st.data_editor(
            df_today[display_cols],
            column_config={
                "MUST": st.column_config.CheckboxColumn("MUST (First)", help="最初に訪問する", default=False),
                "code": "コード",
                "name": "顧客名",
                "sales": st.column_config.NumberColumn("売上見込", format="¥%d"),
                "WorkMinutes": st.column_config.NumberColumn("作業時間(分)", min_value=1, step=1, help="作業時間を編集できます"),
                "NoEntryTime": "入場不可",
                "address": "住所"
            },
            disabled=["code", "name", "sales", "NoEntryTime", "address", "lat", "lng"],
            hide_index=True,
            use_container_width=True,
            key="today_editor"
        )
        
        # 編集結果をsession_stateに反映
        # 行の削除等はdata_editorでは標準で「削除」機能があるが、ここでは編集結果をそのままリストに戻す
        # 注意: 削除機能有効化には num_rows="dynamic" が必要
        
        # data_editorの結果はDataFrameなので、辞書リストに戻してsession_stateを更新
        # ただしrerunループを防ぐため、比較する？ -> data_editorは編集時にrerunするので、ここで代入してOK
        st.session_state['today_list'] = edited_df.to_dict('records')

        # 削除ボタン（一括削除など）
        if st.button("全クリア"):
            st.session_state['today_list'] = []
            st.rerun()

# アクションエリア
st.markdown("---")
st.header("アクション")

col_a, col_b = st.columns(2)

with col_a:
    if st.button("自動並び替え (距離順)", type="primary"):
        if not st.session_state['today_list']:
            st.warning("TODAYリストが空です。")
        else:
            with st.spinner("ルート計算中..."):
                # 起点の座標取得（簡易的に固定値あるいはAPIでジオコーディングが必要）
                # 今回はMVPなので設定ファイルのデフォルト住所に対応する座標をハードコード、あるいはAPIがあるならAPIを使う
                # ここでは「千葉県市原市白金町1-32」の座標を一時的に使用（サンプルに合わせる）
                # 35.534222, 140.111557 (サンプル参照) -> 実際には住所から取るべきだが
                origin_lat, origin_lng = 35.534222, 140.111557 # 仮
                
                # ルート最適化ロジック呼び出し
                # locationsリスト作成 (index 0 は起点)
                locations = [{'lat': origin_lat, 'lng': origin_lng}] + \
                            [{'lat': item['lat'], 'lng': item['lng']} for item in st.session_state['today_list']]
                
                # 距離行列
                dist_matrix, _ = get_distance_matrix(locations, api_key=api_key)
                
                # MUSTフラグが立っている箇所のインデックスを取得
                # locations[0] は起点なので、locations[i+1] が today_list[i] に対応
                # optimize_route に渡す must_visit_indices は locations のインデックス（1オリジン）
                must_indices = []
                for idx, item in enumerate(st.session_state['today_list']):
                    if item.get('MUST', False):
                        must_indices.append(idx + 1) # locationsにおけるインデックス
                
                # 最適化（2-opt）
                # route_indicesは locations のインデックス（1オリジン、0は起点）
                optimized_indices = optimize_route(locations, dist_matrix, must_visit_indices=must_indices)
                
                # 結果をTODAYリストに反映
                # optimized_indices は [3, 1, 2, ...] のような順序（locationsのインデックス）
                # これを today_list のインデックス（0開始）に変換 -> index - 1
                new_today_list = [st.session_state['today_list'][i-1] for i in optimized_indices]
                st.session_state['today_list'] = new_today_list
                st.success("最短ルート順に並び替えました！")
                st.rerun()

with col_b:
    if st.button("訪問予定表 (Excel) 作成"):
        if not st.session_state['today_list']:
            st.warning("リストが空です")
        else:
            # スケジュール計算
            # ここでも起点は仮
            origin_lat, origin_lng = 35.534222, 140.111557
            
            # 並び替え済みのリストを使用
            # インデックスのリスト（0, 1, 2...）を渡す
            indices = range(len(st.session_state['today_list']))
            df_today = pd.DataFrame(st.session_state['today_list'])
            
            schedule = calculate_schedule(
                indices, df_today, 
                origin_lat, origin_lng, 
                departure_time_str.strftime("%H:%M"),
                work_minutes_def,
                lunch_start.strftime("%H:%M"),
                lunch_end.strftime("%H:%M")
            )
            
            # Excel生成
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "VisitPlan"
            
            headers = ["対象日付", "順番", "顧客コード", "顧客名", "住所", "作業時間(分)", "到着時刻", "終了時刻", 
                       "移動時間(分)", "移動距離(km)", "売上見込(円)", "メモ", "GoogleMapURL"]
            ws.append(headers)
            
            today_str = datetime.now().strftime('%Y-%m-%d')
            
            # ルートURL生成ロジック（RouteMapURL）は削除
            
            # for item in schedule:
            #     waypoints_list...
            

            for item in schedule:
                # 距離はメートル -> km
                dist_km = item['travel_dist']
                
                row = [
                    today_str,
                    item['seq'],
                    item['code'],
                    item['name'],
                    item['address'],
                    item['work_min'],
                    item['arrival_time'].strftime('%H:%M'),
                    item['finish_time'].strftime('%H:%M'),
                    item['travel_min'],
                    dist_km,
                    item['sales'],
                    "", # メモ
                    f"https://www.google.com/maps/search/?api=1&query={item['lat']},{item['lng']}", # GoogleMapURL
                    # full_route_url # RouteMapURL (削除)
                ]
                ws.append(row)

            # バイト列に保存
            output = io.BytesIO()
            wb.save(output)
            processed_data = output.getvalue()
            
            st.download_button(
                label="Excelダウンロード",
                data=processed_data,
                file_name=f"VisitPlan_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

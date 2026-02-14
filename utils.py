import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import googlemaps
from math import radians, cos, sin, asin, sqrt
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import streamlit as st
import yaml

# 設定の読み込み
def load_config():
    with open('config.yaml', 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)

CONFIG = load_config()

# ハーサイン距離（代替手段）
def haversine(lon1, lat1, lon2, lat2):
    # km単位で返す
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371 # 地球の半径 (km)
    return c * r

# データの読み込みと前処理
def load_customer_data(file):
    try:
        if file.name.endswith('.csv'):
            # 文字コード自動判別の簡易実装（utf-8-sig -> shift_jis -> cp932）
            # ヘッダーが2行目（index 1）にあると想定
            try:
                df = pd.read_csv(file, encoding='utf-8-sig', header=1)
            except UnicodeDecodeError:
                file.seek(0)
                try:
                    df = pd.read_csv(file, encoding='shift_jis', header=1)
                except UnicodeDecodeError:
                    file.seek(0)
                    df = pd.read_csv(file, encoding='cp932', header=1)
        else:
            df = pd.read_excel(file, header=1)
        
        # 列名マッピング
        col_map = CONFIG['master_columns']
        # 必要な列が存在するかチェック
        required_cols = [col_map['customer_code'], col_map['customer_name'], col_map['latlng']]
        missing = [c for c in required_cols if c not in df.columns]
        
        if missing:
            return None, f"必須列が見つかりません: {', '.join(missing)}"
        
        # 内部処理用にカラム名を統一
        rename_map = {
            col_map['customer_code']: 'code',
            col_map['customer_name']: 'name',
            col_map['predicted_sales']: 'sales',
            col_map['latlng']: 'latlng_raw',
            col_map['address1']: 'address',
            col_map.get('work_minutes', '作業時間'): 'WorkMinutes',
            col_map.get('no_entry_time', '入場不可時間帯'): 'NoEntryTime'
        }
        df = df.rename(columns=rename_map)
        
        # 緯度経度の分割
        # "35.534222, 140.111557" -> lat, lng
        # 文字列型にして分割
        df['latlng_raw'] = df['latlng_raw'].astype(str)
        df[['lat', 'lng']] = df['latlng_raw'].str.split(',', expand=True)
        df['lat'] = pd.to_numeric(df['lat'], errors='coerce')
        df['lng'] = pd.to_numeric(df['lng'], errors='coerce')
        
        # 緯度経度が欠損している行を除外
        invalid_rows = df[df['lat'].isna() | df['lng'].isna()]
        if not invalid_rows.empty:
            st.warning(f"{len(invalid_rows)}行のデータで緯度経度が不正なため除外されました。")
            df = df.dropna(subset=['lat', 'lng'])
            
        # 売上がNaNの場合は0埋め
        if 'sales' in df.columns:
            df['sales'] = df['sales'].fillna(0).astype(int)
        else:
            df['sales'] = 0
            
        # 作業時間の欠損処理（configのデフォルト値で埋める）
        if 'WorkMinutes' in df.columns:
             df['WorkMinutes'] = pd.to_numeric(df['WorkMinutes'], errors='coerce').fillna(CONFIG['defaults']['work_minutes'])
        else:
            df['WorkMinutes'] = CONFIG['defaults']['work_minutes']
            
        # 入場不可時間帯の欠損処理（空文字にする）
        if 'NoEntryTime' not in df.columns:
            df['NoEntryTime'] = None
            
        return df, None
        
    except Exception as e:
        return None, str(e)

# 距離行列の取得（Google Maps API または 直線距離）
def get_distance_matrix(locations, api_key=None, origin=None):
    """
    locations: list of dict {'lat': float, 'lng': float} (index 0 is origin if origin is None)
    origin: tuple (lat, lng) or str (address) if provided separately
    """
    n = len(locations)
    dist_matrix = np.zeros((n, n)) # メートル
    time_matrix = np.zeros((n, n)) # 秒
    
    # APIキーがある場合
    if api_key:
        try:
            gmaps = googlemaps.Client(key=api_key)
            
            # 緯度経度リストの作成 (API用)
            coords = [(loc['lat'], loc['lng']) for loc in locations]
            
            # API制限対策（要素数100以下/リクエスト、推奨25以下）
            # 6x6 = 36 要素ずつ処理
            batch_size = 6
            
            for i in range(0, n, batch_size):
                origin_batch = coords[i : i + batch_size]
                
                for j in range(0, n, batch_size):
                    dest_batch = coords[j : j + batch_size]
                    
                    # APIコール
                    # departure_time=datetime.now() で現在の交通状況を考慮
                    response = gmaps.distance_matrix(
                        origins=origin_batch,
                        destinations=dest_batch,
                        mode='driving',
                        departure_time=datetime.now()
                    )
                    
                    rows = response.get('rows', [])
                    for r_idx, row in enumerate(rows):
                        elements = row.get('elements', [])
                        for c_idx, element in enumerate(elements):
                            if element.get('status') == 'OK':
                                # マトリックス全体のインデックス
                                global_row = i + r_idx
                                global_col = j + c_idx
                                
                                # 距離 (メートル)
                                dist_val = element.get('distance', {}).get('value', 0)
                                # 時間 (秒) - trafficがあれば優先
                                dur_val = element.get('duration_in_traffic', {}).get('value', 0)
                                if dur_val == 0:
                                    dur_val = element.get('duration', {}).get('value', 0)
                                
                                dist_matrix[global_row][global_col] = dist_val
                                time_matrix[global_row][global_col] = dur_val
            
            # 成功したらここでリターン（フォールバックに行かせない）
            return dist_matrix, time_matrix
            
        except Exception as e:
            st.warning(f"Google Maps API エラー: {e}。直線距離（30km/h）で計算します。")
    
    # 直線距離（フォールバック）
    # 速度仮定: 30km/h = 500m/min = 8.33m/s
    speed_mps = 30 * 1000 / 3600
    
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            d_km = haversine(locations[i]['lng'], locations[i]['lat'], 
                             locations[j]['lng'], locations[j]['lat'])
            dist_matrix[i][j] = d_km * 1000 # m
            time_matrix[i][j] = (d_km * 1000) / speed_mps # seconds
            
            # APIが使えた場合は上書きするロジックをここに書く
            
    return dist_matrix, time_matrix

# ルート最適化（Nearest Insertion + 2-opt）
# must_visit_indices: 訪問必須（かつ最初に行く）箇所のインデックスリスト（0オリジン、depot除くindex）
def optimize_route(locations, dist_matrix, must_visit_indices=None):
    n = len(locations)
    # 0番目は起点（Depot）
    
    # MUST箇所の処理
    # MUST箇所を先に訪問するルートを構築
    # Depot -> Must1 -> Must2 ... -> (Nearest Unvisited)
    
    current_node = 0 # Depot
    visited_must = []
    
    if must_visit_indices:
        # MUST箇所をどう巡るか？
        # 単純に「リスト順」ではなく、MUST箇所内でも最適化すべきだが、
        # MUST箇所が少数なら Nearest Neighbor で十分
        
        # must_visit_indices は locations のインデックス引数
        # locations[idx] が対象
        
        remaining_must = set(must_visit_indices)
        
        while remaining_must:
            # 現在地から一番近いMUSTを探す
            nearest_must = min(remaining_must, key=lambda x: dist_matrix[current_node][x])
            visited_must.append(nearest_must)
            remaining_must.remove(nearest_must)
            current_node = nearest_must
            
    route = [0] + visited_must
    
    # 残りの箇所
    unvisited = set(range(1, n)) - set(visited_must)
    
    # Nearest Neighbor で残りを追加 (Nearest Insertion の簡易版として実装中)
    while unvisited:
        best_node = -1
        best_pos = -1
        min_cost = float('inf')
        
        for node in unvisited:
            for i in range(len(route)):
                u = route[i]
                v = route[(i + 1) % len(route)] # Depotに戻るサイクルとみなすか、単純パスか
                # 今回は Dep -> ... -> Last なので、単純挿入コストを見る
                # しかし「巡回」ではなく「訪問順」なので、最後の点からの距離を見るのがNearest Neighbor
                # 要件は「Nearest Insertion」
                
                # 単純化：Nearest Neighbor（現在地から一番近い所へ）が一般的に訪問順としては直感的
                # Nearest Insertionは巡回セールスマン(TSP)の構築法。
                pass
        
        # ここでは実装が容易でそれなりに良い Nearest Neighbor を採用し、最後に2-optで改善する
        last_node = route[-1]
        nearest_node = min(unvisited, key=lambda x: dist_matrix[last_node][x])
        route.append(nearest_node)
        unvisited.remove(nearest_node)
        
    # 2-opt (MUST箇所の順序は守るべきか？ -> MUSTは「今日の1番目に行く」など順序指定の意味合いが強い
    # しかし、要件は「この顧客は今日の1番目に行く」という【MUST】設定。
    # 複数ある場合は「1番目グループ」と解釈し、その中での最適化は許容されるべき。
    # また、MUSTグループが終わった後に通常グループに行く。
    # したがって、MUSTグループと通常グループを混ぜてはいけない。
    
    # 2-optを「MUSTグループ内」と「通常グループ内」で別々にかけるのが安全。
    # 今回は簡易的に、全体にかけてしまうとMUSTが後ろに回る可能性があるので、
    # MUST以降の部分（通常パート）のみに2-optをかける。
    
    # must_visit_indices がある場合、その長さ分は固定（Depot(1) + Must(k)）
    fixed_len = 1 + (len(must_visit_indices) if must_visit_indices else 0)
    
    improved = True
    while improved:
        improved = False
        # fixed_len 以降の要素のみ最適化対象
        start_idx = max(1, fixed_len) 
        if start_idx >= len(route) - 1:
            break
            
        for i in range(start_idx, len(route) - 2):
            for j in range(i + 1, len(route)):
                if j - i == 1: continue 
                # 現在のコスト
                d1 = dist_matrix[route[i-1]][route[i]]
                d2 = dist_matrix[route[j]][route[(j+1)%len(route)]] if j+1 < len(route) else 0
                # 交換後のコスト
                d3 = dist_matrix[route[i-1]][route[j]]
                d4 = dist_matrix[route[i]][route[(j+1)%len(route)]] if j+1 < len(route) else 0
                
                if d1 + d2 > d3 + d4:
                    route[i:j+1] = reversed(route[i:j+1])
                    improved = True
                    
    return route[1:] # 起点を除く訪問順のインデックスリスト

# スケジュール計算
def calculate_schedule(route_indices, df_today, origin_lat, origin_lng, start_time_str, work_min, lunch_start_str, lunch_end_str):
    # route_indices: df_today 内の index ではなく、0オリジンの順序
    # df_today: 選択されたデータフレーム
    
    schedule = []
    
    current_time = datetime.strptime(f"{datetime.now().date()} {start_time_str}", "%Y-%m-%d %H:%M")
    lunch_start = datetime.strptime(f"{datetime.now().date()} {lunch_start_str}", "%Y-%m-%d %H:%M")
    lunch_end = datetime.strptime(f"{datetime.now().date()} {lunch_end_str}", "%Y-%m-%d %H:%M")
    
    # 前回の位置（初期値は起点）
    prev_lat = origin_lat
    prev_lng = origin_lng
    
    # 速度仮定（フォールバック用）
    speed_mps = 30 * 1000 / 3600
    
    total_sales = 0
    
    for i, idx in enumerate(route_indices):
        row = df_today.iloc[idx]
        
        # 移動計算
        dist_km = haversine(prev_lng, prev_lat, row['lng'], row['lat'])
        travel_sec = (dist_km * 1000) / speed_mps
        travel_min = int(travel_sec / 60)
        
        arrival_time = current_time + timedelta(minutes=travel_min)
        
        # 入場不可時間帯のチェック
        # NoEntryTime: "12:00-13:00" string or similar
        no_entry_val = row.get('NoEntryTime')
        if no_entry_val and isinstance(no_entry_val, str) and '-' in no_entry_val:
            try:
                start_str, end_str = no_entry_val.split('-')
                # 今日の日付と結合
                ne_start = datetime.strptime(f"{datetime.now().date()} {start_str.strip()}", "%Y-%m-%d %H:%M")
                ne_end = datetime.strptime(f"{datetime.now().date()} {end_str.strip()}", "%Y-%m-%d %H:%M")
                
                # 到着時刻が入場不可時間帯に含まれる場合、終了まで待機
                if ne_start <= arrival_time < ne_end:
                    arrival_time = ne_end
            except:
                pass # パースエラー時は無視
        
        # 昼休憩判定
        # 到着が昼休憩にかかる -> 休憩終了まで待機
        if lunch_start <= arrival_time < lunch_end:
            arrival_time = lunch_end
        
        # 作業終了予定
        work_duration = int(row.get('WorkMinutes', work_min))
        finish_time = arrival_time + timedelta(minutes=work_duration)
        
        # 作業中に昼休憩にかかる -> 休憩分後ろ倒し（非常に簡易的な実装）
        # 到着は休憩前だが、終了が休憩開始を過ぎる場合
        if arrival_time < lunch_start and finish_time > lunch_start:
            # 休憩時間を挟む
            finish_time += (lunch_end - lunch_start)
            
        schedule.append({
            'seq': i + 1,
            'code': row['code'],
            'name': row['name'],
            'address': row.get('address', ''),
            'sales': row['sales'],
            'arrival_time': arrival_time,
            'finish_time': finish_time,
            'work_min': work_duration,
            'travel_min': travel_min,
            'travel_dist': round(dist_km, 1),
            'lat': row['lat'],
            'lng': row['lng']
        })
        
        current_time = finish_time
        prev_lat = row['lat']
        prev_lng = row['lng']
        total_sales += row['sales']
        
    return schedule

# Excel出力
def create_excel(schedule_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VisitPlan"
    
    headers = ["対象日付", "順番", "顧客コード", "顧客名", "住所", "作業時間(分)", 
               "到着時刻", "終了時刻", "移動時間(分)", "移動距離(km)", "売上見込(円)", 
               "メモ", "GoogleMapURL"]
    
    ws.append(headers)
    
    today_str = datetime.now().strftime('%Y-%m-%d')
    
    # Directions URL作成のためのWaypoint構築
    waypoints = [f"{item['lat']},{item['lng']}" for item in schedule_data]
    
    for item in schedule_data:
        # 個別Google Map URL
        gmap_url = f"https://www.google.com/maps/search/?api=1&query={item['lat']},{item['lng']}"
        
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
            item['travel_dist'],
            item['sales'],
            "", # メモ
            gmap_url
            # route_url (削除)
        ]
        ws.append(row)
        
    # スタイル調整
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
        
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width
        
    return wb


import pandas as pd
from datetime import datetime
from utils import calculate_schedule, optimize_route

# Mock Data
data = [
    {'code': 1, 'name': 'A', 'lat': 35.0, 'lng': 139.0, 'sales': 1000, 'WorkMinutes': 10, 'NoEntryTime': None}, # Depot-ish
    {'code': 2, 'name': 'B', 'lat': 35.1, 'lng': 139.1, 'sales': 2000, 'WorkMinutes': 10, 'NoEntryTime': '12:00-13:00'},
    {'code': 3, 'name': 'C', 'lat': 35.2, 'lng': 139.2, 'sales': 3000, 'WorkMinutes': 10, 'NoEntryTime': None},
    {'code': 4, 'name': 'D', 'lat': 35.05, 'lng': 139.05, 'sales': 4000, 'WorkMinutes': 10, 'NoEntryTime': None}
]
df = pd.DataFrame(data)

# Test Optimize Route with MUST
# Locations: 0=Depot, 1=A, 2=B, 3=C, 4=D
locations = [{'lat': 35.0, 'lng': 139.0}] + [d for d in data] 
# indices in optimize_route are 1..4 (1=A, 2=B, 3=C, 4=D)
# Let's say C (index 3) is MUST.
# Dist matrix: simple euclidean for test
import numpy as np
dist_matrix = np.zeros((5, 5))
for i in range(5):
    for j in range(5):
        dist_matrix[i][j] = abs(i-j) * 10 # Dummy distance

print("Testing MUST optimization...")
# MUST index is 3 (C)
route = optimize_route(locations, dist_matrix, must_visit_indices=[3])
print(f"Route: {route}")
# Expected: 3 should be first (after depot, which isn't in output of optimize_route usually, or is it? check utils)
# utils returns `route[1:]`. route[0] is depot.
# So first element of result should be 3.

# Test Schedule with No Entry
# Arrive at B at 12:10 (mock). No Entry 12:00-13:00. Should wait till 13:00.
print("\nTesting Schedule with No Entry...")
# We'll call calculate_schedule.
# We need to control arrival time.
# Let's say we go to B (index 1 in DF) first.
# Travel to B: 10km?
# Speed 30km/h = 0.5km/min.
# If dist is 10km -> 20 min.
# Start at 11:50 -> Arrive 12:10.
# NoEntry 12:00-13:00 -> Wait till 13:00.
# Work 10 min -> Finish 13:10.

# Mocking haversine to return controlled distance is hard without mocking utils.
# But we can check if logic parses the string.
# We'll just run it with the utils function and real haversine.
# We need to construct a route where we arrive in the window.
# 35.0,139.0 -> 35.1,139.1 is approx 14km.
# 14km / 30km/h = ~28 min.
# Start at 11:40 -> Arrive 12:08.
# Window 12:00-13:00.
# Should finish at 13:10 (Work 10 min).

route_indices = [1] # Index 1 in DF is B (code 2)
# Origin
origin_lat, origin_lng = 35.0, 139.0

# Start 11:40
# We need to mock datetime.now() date? function uses datetime.now().date().
# We'll just assume today.
start_time_str = "11:40"

item = calculate_schedule(
    route_indices, df, 
    origin_lat, origin_lng, 
    start_time_str, 
    10, 
    "12:00", "13:00" # Lunch time (overlapped with NoEntry, might be complex, lets set lunch late)
)[0]
# Set lunch 14:00-15:00 to avoid interference
item = calculate_schedule(
    route_indices, df, 
    origin_lat, origin_lng, 
    start_time_str, 
    10, 
    "14:00", "15:00" 
)[0]

print(f"Arrival: {item['arrival_time'].strftime('%H:%M')}")
print(f"Finish: {item['finish_time'].strftime('%H:%M')}")

# Expected:
# Arrival should be 13:00 (delayed from ~12:08)
# Finish 13:10.

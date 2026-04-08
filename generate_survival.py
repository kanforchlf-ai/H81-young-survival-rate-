"""
生成存活率分析資料：青年大區大專，2023年後新加入成員
存活定義：尚未出現「連續4週未出席」事件
"""
import xlrd
import json
from datetime import date

XLS_PATH = '/Users/nn/Desktop/claude-inbox/81主日歷史資料.xls'

wb = xlrd.open_workbook(XLS_PATH, ignore_workbook_corruption=True)
ws = wb.sheets()[0]

# ── 建立欄位 -> 大約日期的對應 ──────────────────────────────
WEEK_OFFSET = {
    '第一週': 0, '第二週': 7, '第三週': 14,
    '第四週': 21, '第五週': 28,
}
MONTH_MAP = {
    '1月':1,'2月':2,'3月':3,'4月':4,'5月':5,'6月':6,
    '7月':7,'8月':8,'9月':9,'10月':10,'11月':11,'12月':12,
}

col_dates = []       # index i => date object for column (8+i)
cur_year_month = None
for c in range(8, ws.ncols):
    v0 = str(ws.cell_value(0, c)).strip()
    v1 = str(ws.cell_value(1, c)).strip()
    if v0:
        # parse "2022年4月"
        parts = v0.replace('年', ' ').replace('月', '').split()
        yr, mo = int(parts[0]), int(parts[1])
        cur_year_month = (yr, mo)
    if cur_year_month:
        yr, mo = cur_year_month
        day = 1 + WEEK_OFFSET.get(v1, 0)
        try:
            d = date(yr, mo, min(day, 28))
        except:
            d = date(yr, mo, 1)
        col_dates.append(d)
    else:
        col_dates.append(None)

NUM_WEEKS = len(col_dates)

# ── 讀取成員資料 ────────────────────────────────────────────
YOUNG_ZONE = '青年大區'
TARGET_GROUP = '大專'
CUTOFF_YEAR = 2023

members = []
for r in range(2, ws.nrows):
    zone = str(ws.cell_value(r, 0)).strip()
    sub_zone = str(ws.cell_value(r, 1)).strip()
    team = str(ws.cell_value(r, 2)).strip()
    name = str(ws.cell_value(r, 3)).strip()
    group = str(ws.cell_value(r, 5)).strip()
    baptism = str(ws.cell_value(r, 7)).strip()

    if zone != YOUNG_ZONE or group != TARGET_GROUP:
        continue

    # 取出逐週出席（0/1）
    attend = []
    for c in range(8, ws.ncols):
        try:
            v = float(ws.cell_value(r, c))
            attend.append(1 if v >= 1 else 0)
        except:
            attend.append(0)

    # 找首次出席 col index
    first_idx = next((i for i, v in enumerate(attend) if v == 1), None)
    if first_idx is None:
        continue

    first_date = col_dates[first_idx]
    if first_date is None:
        continue

    # 判斷是否為「新加入」（2023年後）
    first_year = first_date.year
    baptism_year = None
    if baptism and len(baptism) >= 4 and baptism[:4].isdigit():
        baptism_year = int(baptism[:4])

    is_new = (first_year >= CUTOFF_YEAR) or (baptism_year is not None and baptism_year >= CUTOFF_YEAR)
    if not is_new:
        continue

    join_year = str(first_year)
    join_ym = f"{first_date.year}年{first_date.month}月"

    # 從首次出席起的出席序列
    seq = attend[first_idx:]
    total_weeks = len(seq)

    # ── 存活事件：連續4週未出席 ─────────────────────────────
    # survival_week = 觸發事件時距首次出席的週數（None=尚未發生）
    survival_week = None
    consec = 0
    for w, v in enumerate(seq):
        if v == 0:
            consec += 1
            if consec >= 4:
                # 事件觸發點 = 連續缺席開始的那週
                survival_week = w - 3
                break
        else:
            consec = 0

    # ── 逐月出席率（每4週一組）──────────────────────────────
    monthly_rate = []
    for m in range(0, min(total_weeks, 60), 4):
        chunk = seq[m:m+4]
        if len(chunk) == 0:
            break
        rate = sum(chunk) / len(chunk)
        monthly_rate.append(round(rate * 100, 1))

    # ── 整體出席率 ──────────────────────────────────────────
    usable = min(total_weeks, 52)
    attend_count = sum(seq[:usable])
    attend_rate = round(attend_count / usable * 100, 1)

    members.append({
        '小區': sub_zone,
        '排': team,
        '姓名': name,
        '加入年': join_year,
        '首次出席': join_ym,
        '總週數': total_weeks,
        '出席數': attend_count,
        '出席率': attend_rate,
        '流失週': survival_week,   # None=仍存活/無資料 / int=第幾週觸發事件
        '每月出席率': monthly_rate,
    })

print(f'符合條件成員：{len(members)} 人')
for y in ['2023','2024','2025']:
    cnt = sum(1 for m in members if m['加入年'] == y)
    lost = sum(1 for m in members if m['加入年'] == y and m['流失週'] is not None)
    print(f'  {y}: {cnt} 人，已流失 {lost} 人')

# ── 計算存活曲線 ────────────────────────────────────────────
# 對每個時間點 t (週)，還存活的比例 = 尚未觸發流失事件且有該週資料的人
MAX_WEEKS = 72

def survival_curve(group_members, max_w=MAX_WEEKS):
    curve = []
    n = len(group_members)
    if n == 0:
        return []
    for t in range(0, max_w + 1, 2):
        # 在週 t 仍有資料且尚未流失的人數
        at_risk = sum(1 for m in group_members if m['總週數'] > t)
        survived = sum(1 for m in group_members
                       if m['總週數'] > t and
                       (m['流失週'] is None or m['流失週'] > t))
        if at_risk == 0:
            break
        curve.append({
            'week': t,
            'pct': round(survived / at_risk * 100, 1),
            'at_risk': at_risk,
            'survived': survived,
        })
    return curve

all_curve = survival_curve(members)
curves_by_year = {}
for y in ['2023', '2024', '2025']:
    g = [m for m in members if m['加入年'] == y]
    curves_by_year[y] = survival_curve(g)

# ── 流失時機分佈（histogram，每8週一組）──────────────────────
dropout_dist = {}
for m in members:
    if m['流失週'] is not None:
        bucket = (m['流失週'] // 8) * 8
        key = f"{bucket}-{bucket+7}週"
        dropout_dist[key] = dropout_dist.get(key, 0) + 1

# ── 每週平均出席率（對有資料的人取平均）────────────────────
MAX_TREND = 60
avg_weekly = []
for t in range(MAX_TREND):
    vals = [m['每月出席率'][t // 4] for m in members
            if t // 4 < len(m['每月出席率'])]
    if not vals:
        break
    avg_weekly.append({'week': t, 'avg': round(sum(vals) / len(vals), 1), 'n': len(vals)})

# ── 倖存者（觸發事件 > 8週，或從未觸發）────────────────────
survivors = [m for m in members if m['流失週'] is None or m['流失週'] > 8]
print(f'\n倖存者（8週以上）: {len(survivors)} 人')
for s in survivors:
    print(f"  {s['姓名']} ({s['加入年']}) dropout@{s['流失週']} rate={s['出席率']}% trend={s['每月出席率'][:12]}")

# ── 輸出 ────────────────────────────────────────────────────
output = {
    'members': members,
    'survival_curve': all_curve,
    'survival_by_year': curves_by_year,
    'dropout_dist': dropout_dist,
    'avg_weekly_trend': avg_weekly,
    'survivors': survivors,
    'summary': {
        'total': len(members),
        'still_active': sum(1 for m in members if m['流失週'] is None),
        'dropped': sum(1 for m in members if m['流失週'] is not None),
        'by_year': {
            y: {
                'total': sum(1 for m in members if m['加入年'] == y),
                'active': sum(1 for m in members if m['加入年'] == y and m['流失週'] is None),
                'dropped': sum(1 for m in members if m['加入年'] == y and m['流失週'] is not None),
            }
            for y in ['2023', '2024', '2025']
        }
    }
}

with open('/tmp/h81-viz/survival_data.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print('已輸出 survival_data.json')
print('Summary:', json.dumps(output['summary'], ensure_ascii=False))
print('流失時機分佈:', dropout_dist)

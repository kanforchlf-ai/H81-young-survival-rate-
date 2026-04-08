#!/usr/bin/env python3
"""
add_new_grade_members.py
把 藍聿昕、林沛均 加入 grade_analysis.json 大四
"""
import xlrd, json
from datetime import date

XLS_PATH = '/Users/nn/Desktop/claude-inbox/81主日歷史資料.xls'
TARGETS = {'藍聿昕': '大四', '林沛均': '大四'}

wb = xlrd.open_workbook(XLS_PATH, ignore_workbook_corruption=True)
ws = wb.sheets()[0]

WEEK_OFFSET = {'第一週':0,'第二週':7,'第三週':14,'第四週':21,'第五週':28}
col_dates = []
cur_ym = None
for c in range(8, ws.ncols):
    v0 = str(ws.cell_value(0,c)).strip()
    v1 = str(ws.cell_value(1,c)).strip()
    if v0:
        parts = v0.replace('年',' ').replace('月','').split()
        cur_ym = (int(parts[0]), int(parts[1]))
    if cur_ym:
        yr, mo = cur_ym
        day = 1 + WEEK_OFFSET.get(v1, 0)
        try: d = date(yr, mo, min(day, 28))
        except: d = date(yr, mo, 1)
        col_dates.append(d)
    else:
        col_dates.append(None)

found = {}
for r in range(2, ws.nrows):
    zone  = str(ws.cell_value(r,0)).strip()
    group = str(ws.cell_value(r,5)).strip()
    name  = str(ws.cell_value(r,3)).strip()
    team  = str(ws.cell_value(r,2)).strip()
    if zone != '青年大區' or group != '大專': continue
    if name not in TARGETS: continue

    attend = []
    for c in range(8, ws.ncols):
        try: v = float(ws.cell_value(r,c))
        except: v = 0
        attend.append(1 if v>=1 else 0)

    first_idx = next((i for i,v in enumerate(attend) if v==1), None)
    if first_idx is None: continue
    first_date = col_dates[first_idx]
    join_year = str(first_date.year)
    seq = attend[first_idx:]
    total_weeks = len(seq)

    # dropout_week
    consec = 0; dropout_week = None
    for w, v in enumerate(seq):
        if v == 0:
            consec += 1
            if consec >= 4: dropout_week = w - 3; break
        else: consec = 0

    usable = min(total_weeks, 52)
    attend_rate = round(sum(seq[:usable]) / usable * 100, 1)
    recent_seq = seq[-26:] if len(seq) >= 26 else seq
    recent_rate = round(sum(recent_seq) / len(recent_seq) * 100, 1)

    found[name] = {
        'name': name,
        'grade': TARGETS[name],
        'total_weeks': total_weeks,
        'attend_rate': attend_rate,
        'recent_rate': recent_rate,
        'dropout_week': dropout_week,
        'join_year': join_year,
    }
    print(f"找到 {name}: join={join_year}, dropout={dropout_week}, attend={attend_rate}%, recent={recent_rate}%")

if len(found) != len(TARGETS):
    print("警告：未找到", set(TARGETS) - set(found))

# 加入 grade_analysis.json 大四
with open('grade_analysis.json') as f:
    grade_data = json.load(f)

for name, member in found.items():
    g = member['grade']
    # 避免重複
    existing = [m for m in grade_data[g]['members'] if m['name'] == name]
    if existing:
        print(f"{name} 已存在，跳過")
        continue
    grade_data[g]['members'].append(member)

# 重算大四統計
for g in ['大四']:
    ms = grade_data[g]['members']
    n = len(ms)
    dropped = sum(1 for m in ms if m['dropout_week'] is not None)
    grade_data[g]['n'] = n
    grade_data[g]['dropped'] = dropped
    grade_data[g]['never_dropped'] = n - dropped
    grade_data[g]['avg_attend_rate'] = round(sum(m['attend_rate'] for m in ms)/n, 1)
    grade_data[g]['avg_recent_rate'] = round(sum(m['recent_rate'] for m in ms)/n, 1)
    from collections import defaultdict
    by_year = defaultdict(int)
    for m in ms:
        by_year[m['join_year']] += 1
    grade_data[g]['by_join_year'] = dict(by_year)
    print(f"大四更新後：n={n}, dropped={dropped}")

with open('grade_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(grade_data, f, ensure_ascii=False, indent=2)

print("grade_analysis.json 已更新")

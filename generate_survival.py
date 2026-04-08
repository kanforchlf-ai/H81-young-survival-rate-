"""
生成存活率分析資料：青年大區大專，全部 185 人
存活定義：尚未出現「連續 4 週未出席」事件
"""
import xlrd
import json
from datetime import date

XLS_PATH = '/Users/nn/Desktop/claude-inbox/81主日歷史資料.xls'

wb = xlrd.open_workbook(XLS_PATH, ignore_workbook_corruption=True)
ws = wb.sheets()[0]

# ── 欄位日期對應 ─────────────────────────────────────────────
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

# ── 讀取所有青年大區大專成員 ─────────────────────────────────
members = []
for r in range(2, ws.nrows):
    zone  = str(ws.cell_value(r,0)).strip()
    sub   = str(ws.cell_value(r,1)).strip()
    team  = str(ws.cell_value(r,2)).strip()
    name  = str(ws.cell_value(r,3)).strip()
    group = str(ws.cell_value(r,5)).strip()
    bapt  = str(ws.cell_value(r,7)).strip()
    if zone != '青年大區' or group != '大專': continue

    attend = []
    for c in range(8, ws.ncols):
        try: v = float(ws.cell_value(r,c))
        except: v = 0
        attend.append(1 if v>=1 else 0)

    if sum(attend) <= 1: continue   # 排除只來過一次的端值
    first_idx = next((i for i,v in enumerate(attend) if v==1), None)
    if first_idx is None: continue
    first_date = col_dates[first_idx]
    if first_date is None: continue

    join_year = str(first_date.year)
    join_ym   = f"{first_date.year}年{first_date.month}月"
    seq = attend[first_idx:]
    total_weeks = len(seq)
    total_attend = sum(seq)

    # ── 流失事件：連續4週缺席 ────────────────────────────────
    consec = 0
    dropout_week = None
    for w, v in enumerate(seq):
        if v == 0:
            consec += 1
            if consec >= 4:
                dropout_week = w - 3
                break
        else:
            consec = 0

    # ── 每月出席率（每4週一組）──────────────────────────────
    monthly_rate = []
    for m in range(0, min(total_weeks, 60), 4):
        chunk = seq[m:m+4]
        if not chunk: break
        monthly_rate.append(round(sum(chunk)/len(chunk)*100, 1))

    usable = min(total_weeks, 52)
    attend_rate = round(sum(seq[:usable]) / usable * 100, 1)

    members.append({
        '小區': sub, '排': team, '姓名': name,
        '加入年': join_year, '首次出席': join_ym,
        '總週數': total_weeks, '出席數': total_attend,
        '出席率': attend_rate,
        '流失週': dropout_week,
        '每月出席率': monthly_rate,
    })

print(f'總人數：{len(members)}')
YEARS = ['2022','2023','2024','2025']
for y in YEARS:
    g = [m for m in members if m['加入年']==y]
    lost = sum(1 for m in g if m['流失週'] is not None)
    print(f'  {y}: {len(g)} 人，流失 {lost}，從未流失 {len(g)-lost}')

# ── 存活曲線 ─────────────────────────────────────────────────
def survival_curve(grp, step=4, max_w=220):
    curve = []
    for t in range(0, max_w+1, step):
        at_risk = [m for m in grp if m['總週數'] > t]
        if not at_risk: break
        survived = [m for m in at_risk
                    if m['流失週'] is None or m['流失週'] > t]
        curve.append({
            'week': t,
            'pct': round(len(survived)/len(at_risk)*100, 1),
            'at_risk': len(at_risk),
            'survived': len(survived),
        })
    return curve

all_curve  = survival_curve(members)
curves_by_year = {y: survival_curve([m for m in members if m['加入年']==y]) for y in YEARS}

# ── 流失時機分佈 ─────────────────────────────────────────────
from collections import Counter
fine_dist = Counter(m['流失週'] for m in members if m['流失週'] is not None)
bucket_dist = {}
for w, cnt in fine_dist.items():
    b = f"{(w//8)*8}-{(w//8)*8+7}週"
    bucket_dist[b] = bucket_dist.get(b, 0) + cnt

# ── 倖存者（從未流失 or 超過52週才流失）────────────────────
survivors_all = sorted(
    [m for m in members if m['流失週'] is None or m['流失週'] > 52],
    key=lambda m: -(m['流失週'] or 9999)
)
print(f'\n長期存活者（超過52週未流失）：{len(survivors_all)} 人')
for s in survivors_all:
    print(f"  {s['姓名']} ({s['加入年']}) dropout@{s['流失週']} rate={s['出席率']}%")

# ── 輸出 ─────────────────────────────────────────────────────
output = {
    'members': members,
    'survival_curve': all_curve,
    'survival_by_year': curves_by_year,
    'bucket_dist': bucket_dist,
    'fine_dist': {str(k): v for k,v in sorted(fine_dist.items())},
    'survivors_longterm': survivors_all,
    'summary': {
        'total': len(members),
        'never_dropped': sum(1 for m in members if m['流失週'] is None),
        'by_year': {
            y: {
                'total': sum(1 for m in members if m['加入年']==y),
                'never_dropped': sum(1 for m in members if m['加入年']==y and m['流失週'] is None),
                'dropped': sum(1 for m in members if m['加入年']==y and m['流失週'] is not None),
                'week1_dropout': sum(1 for m in members if m['加入年']==y and m['流失週']==1),
            }
            for y in YEARS
        }
    }
}

with open('/tmp/h81-viz/survival_data.json', 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print('\n已輸出 survival_data.json')
print('Summary:', json.dumps(output['summary'], ensure_ascii=False))
print('流失時機分佈:', bucket_dist)

#!/usr/bin/env python3
"""
update_grade_filtered.py
把 grade.html 的資料套用 Option C 篩選：
只保留 survival_data.json 裡出現的穩定成員。
"""
import json, re
from collections import defaultdict

# ── 載入資料 ──────────────────────────────────────────────────
with open('survival_data.json') as f:
    surv = json.load(f)
with open('grade_analysis.json') as f:
    grade_data = json.load(f)

# 51 個穩定成員姓名 set
stable_names = {m['姓名'] for m in surv['members']}
# 姓名 → 成員資料（含近半年出席率）
stable_map = {m['姓名']: m for m in surv['members']}

# ── 用 grade_analysis.json 裡的 recent_rate ────────────────
# grade_analysis.json 的 members[].recent_rate 是近半年出席率
# survival_data.json 的 members[].出席率 是歷史總出席率

grades = ['大一', '大二', '大三', '大四']

def km_curve(members_list):
    """簡易 Kaplan-Meier 曲線（以每個成員的相對週計）"""
    n = len(members_list)
    if n == 0:
        return []
    # 找最大追蹤週數
    max_week = max(m['total_weeks'] for m in members_list)

    # 逐步計算：每 4 週一個點
    events = defaultdict(int)   # dropout_week → 流失人數
    censored = defaultdict(int) # total_weeks → 截斷人數（未流失）

    for m in members_list:
        dw = m['dropout_week']
        tw = m['total_weeks']
        if dw is None:
            censored[tw] += 1
        else:
            events[dw] += 1

    # 在每個 event time 計算 KM
    # 先列出所有 event times
    event_times = sorted(events.keys())

    curve = [{'week': 0, 'pct': 100.0, 'at_risk': n, 'survived': n}]

    survived = n
    at_risk = n

    # 按 4 週步距輸出
    step = 4
    max_out = ((max_week // step) + 1) * step

    prev_pct = 100.0
    cur_at_risk = n

    for w in range(step, max_out + 1, step):
        # 在 [prev_w+1, w] 之間發生的 events
        prev_w = w - step
        for et in event_times:
            if prev_w < et <= w:
                # KM update
                if cur_at_risk > 0:
                    prev_pct = prev_pct * (1 - events[et] / cur_at_risk)
                cur_at_risk -= events[et]
        # censored between prev_w+1 and w (people who finished tracking without event)
        for ct in list(censored.keys()):
            if prev_w < ct <= w:
                cur_at_risk -= censored[ct]

        curve.append({
            'week': w,
            'pct': round(prev_pct, 1),
            'at_risk': cur_at_risk,
            'survived': round(cur_at_risk * prev_pct / 100) if cur_at_risk > 0 else 0
        })

    return curve

# ── 篩選並重算 ────────────────────────────────────────────────
new_grade_data = {}
new_drift_data = {}
total_n = 0

for g in grades:
    orig_members = grade_data[g]['members']
    # 只保留穩定成員
    filtered = [m for m in orig_members if m['name'] in stable_names]

    n = len(filtered)
    total_n += n

    if n == 0:
        new_grade_data[g] = {
            'members': [], 'n': 0, 'dropped': 0, 'never_dropped': 0,
            'avg_attend_rate': 0, 'avg_recent_rate': 0,
            'by_join_year': {}, 'survival_curve': []
        }
        new_drift_data[g] = {'穩定中': 0, '流失': 0, '從未穩定': 0, 'total': 0}
        continue

    dropped = sum(1 for m in filtered if m['dropout_week'] is not None)
    never_dropped = n - dropped
    avg_attend = round(sum(m['attend_rate'] for m in filtered) / n, 1)
    avg_recent = round(sum(m['recent_rate'] for m in filtered) / n, 1)

    by_join = defaultdict(int)
    for m in filtered:
        by_join[m['join_year']] += 1

    # survival curve — 用 dropout_week, total_weeks
    surv_curve = km_curve(filtered)

    new_grade_data[g] = {
        'members': filtered,
        'n': n,
        'dropped': dropped,
        'never_dropped': never_dropped,
        'avg_attend_rate': avg_attend,
        'avg_recent_rate': avg_recent,
        'by_join_year': dict(by_join),
        'survival_curve': surv_curve
    }

    # DRIFT_DATA：所有人都是 Option C 穩定，從未穩定 = 0
    # 穩定中 = recent_rate >= 50%; 流失 = < 50%
    stable_now = sum(1 for m in filtered if m['recent_rate'] >= 50)
    drifted = n - stable_now
    new_drift_data[g] = {
        '穩定中': stable_now,
        '流失': drifted,
        '從未穩定': 0,
        'total': n
    }

grade_summary = ', '.join(g + '=' + str(new_grade_data[g]['n']) for g in grades)
print(f"篩選後各年級人數：{grade_summary}")
print(f"總計：{total_n} 人")
print(f"DRIFT_DATA: {json.dumps(new_drift_data, ensure_ascii=False)}")

# ── 注入 grade.html ───────────────────────────────────────────
with open('grade.html', encoding='utf-8') as f:
    html = f.read()

# 替換 GRADE_DATA
grade_json = json.dumps(new_grade_data, ensure_ascii=False)
html = re.sub(
    r'const GRADE_DATA = \{.*?\};',
    f'const GRADE_DATA = {grade_json};',
    html, flags=re.DOTALL
)

# 替換 DRIFT_DATA
drift_json = json.dumps(new_drift_data, ensure_ascii=False)
html = re.sub(
    r'const DRIFT_DATA = \{.*?\};',
    f'const DRIFT_DATA = {drift_json};',
    html, flags=re.DOTALL
)

# 替換 header 人數
html = re.sub(
    r'共 \d+ 人（已依實際年級校正）',
    f'共 {total_n} 人（Option C 穩定成員）',
    html
)

with open('grade.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("grade.html 已更新")

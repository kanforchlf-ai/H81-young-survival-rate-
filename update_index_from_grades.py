#!/usr/bin/env python3
"""
update_index_from_grades.py
把 index.html 的資料縮限到 grade.html 同樣的 43 人（有年級資料的 Option C 穩定成員）
"""
import json, re

with open('survival_data.json') as f:
    surv = json.load(f)
with open('grade_analysis.json') as f:
    grade_data = json.load(f)

# 有年級資料且排除外地唸書的成員名單
exclude_outside = {m['name'] for g in grade_data.values() for m in g['members']
                   if m.get('reason') == '外地唸書'}
graded_names = {m['name'] for g in grade_data.values() for m in g['members']}

# 從 survival_data.json 篩選（排除外地唸書）
members_43 = [m for m in surv['members']
              if m['姓名'] in graded_names and m['姓名'] not in exclude_outside]
print(f"篩選後：{len(members_43)} 人（grade_analysis 共 {len(graded_names)} 人，其中 {len(members_43)} 在 Option C 名單內）")

# ── 存活曲線（照 generate_survival.py 的邏輯）─────────────────
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

YEARS = ['2022','2023','2024','2025']
all_curve = survival_curve(members_43)
curves_by_year = {y: survival_curve([m for m in members_43 if m['加入年']==y]) for y in YEARS}
survivors_all = sorted(
    [m for m in members_43 if m['流失週'] is None or m['流失週'] > 52],
    key=lambda m: -(m['流失週'] or 9999)
)
print(f"長期存活者：{len(survivors_all)} 人")

# ── 統計 ─────────────────────────────────────────────────────
total = len(members_43)
never_dropped = sum(1 for m in members_43 if m['流失週'] is None)
w1_dropout = sum(1 for m in members_43 if m['流失週'] == 1)
w1_pct = round(w1_dropout / total * 100, 1)

w8_all = next((p for p in all_curve if p['week']==8), None)
w8_pct = w8_all['pct'] if w8_all else 0

# 近期活躍（近8週有出席）
active = sum(1 for m in members_43 if any(r > 0 for r in m['每月出席率'][-2:]))

# Per-year stats
year_stats = {}
for y in YEARS:
    g = [m for m in members_43 if m['加入年']==y]
    n = len(g)
    nd = sum(1 for m in g if m['流失週'] is None)
    w1 = sum(1 for m in g if m['流失週'] == 1)
    c = curves_by_year[y]
    w8 = next((p for p in c if p['week']==8), None)
    w52 = next((p for p in c if p['week']==52), None)
    year_stats[y] = {
        'n': n, 'never_dropped': nd, 'w1': w1,
        'w8_pct': w8['pct'] if w8 else None,
        'w52_pct': w52['pct'] if w52 else None,
    }
    print(f"{y}: n={n}, nd={nd}, w1={w1}, 8wk={w8['pct'] if w8 else '-'}%, 52wk={w52['pct'] if w52 else '-'}%")

# MEMBER_LIST（依出席率排序）
member_list = sorted(members_43, key=lambda m: -m['出席率'])
all_members_js = {
    'attend_rate': round(sum(m['出席率'] for m in members_43)/total, 1),
    'total': total,
    'never_dropped': never_dropped,
}

# ── 注入 index.html ───────────────────────────────────────────
with open('index.html', encoding='utf-8') as f:
    html = f.read()

# Inline JSON replacements
def replace_js_const(html, name, value):
    json_val = json.dumps(value, ensure_ascii=False)
    pattern = rf'const {name} = \[.*?\];'
    new_val = f'const {name} = {json_val};'
    result = re.sub(pattern, new_val, html, flags=re.DOTALL)
    if result == html:
        print(f"警告：未找到 {name}")
    return result

def replace_js_const_obj(html, name, value):
    json_val = json.dumps(value, ensure_ascii=False)
    pattern = rf'const {name} = \{{.*?\}};'
    new_val = f'const {name} = {json_val};'
    result = re.sub(pattern, new_val, html, flags=re.DOTALL)
    if result == html:
        print(f"警告：未找到 {name}")
    return result

html = replace_js_const(html, 'SURVIVAL_CURVE', all_curve)
html = replace_js_const(html, 'CURVE_2022', curves_by_year['2022'])
html = replace_js_const(html, 'CURVE_2023', curves_by_year['2023'])
html = replace_js_const(html, 'CURVE_2024', curves_by_year['2024'])
html = replace_js_const(html, 'CURVE_2025', curves_by_year['2025'])
html = replace_js_const(html, 'SURVIVORS', survivors_all)
html = replace_js_const(html, 'MEMBER_LIST', member_list)
html = replace_js_const_obj(html, 'ALL_MEMBERS', all_members_js)

# Summary cards
html = re.sub(r'<div class="num">\d+</div><div class="label">曾穩定出席（13週≥50%）</div>',
    f'<div class="num">{total}</div><div class="label">曾穩定出席（13週≥50%）</div>', html)
html = re.sub(r'<div class="num danger">[\d.]+%</div><div class="label">第 1 週就流失',
    f'<div class="num danger">{w1_pct}%</div><div class="label">第 1 週就流失', html)
html = re.sub(r'<div class="num warn">[\d.]+%</div><div class="label">全體 8 週存活率',
    f'<div class="num warn">{w8_pct}%</div><div class="label">全體 8 週存活率', html)
html = re.sub(r'<div class="num ok">(\d+) 人</div><div class="label">目前仍活躍',
    f'<div class="num ok">{active} 人</div><div class="label">目前仍活躍', html)

# Insight box
html = re.sub(r'共 \*\*\d+ 人\*\*', f'共 **{total} 人**', html)

# 2022 cohort card
ys = year_stats['2022']
html = re.sub(
    r'(2022年加入（)\d+(人）)',
    f'\\g<1>{ys["n"]}\\g<2>', html)
w1_2022_pct = round(ys['w1']/ys['n']*100) if ys['n'] else 0
html = re.sub(
    r'第1週流失</span><br><strong[^>]*>\d+人（\d+%）</strong>',
    f'第1週流失</span><br><strong style="font-size:1.3rem">{ys["w1"]}人（{w1_2022_pct}%）</strong>',
    html, count=1)
html = re.sub(
    r'(8週存活率</span><br><strong[^>]*>)[\d.]+(%</strong>.*?52週存活率</span><br><strong[^>]*>)[\d.]+(%</strong>.*?從未流失</span><br><strong[^>]*style="[^"]*#2563eb[^"]*">\d+人</strong>)',
    lambda m: m.group(0), html)  # skip - do individually below

# Update 8wk/52wk/never_dropped for each cohort
def update_cohort(html, year, stats):
    # 8週存活率
    w8 = stats['w8_pct']
    w52 = stats['w52_pct']
    nd = stats['never_dropped']
    n = stats['n']
    w1 = stats['w1']
    w1p = round(w1/n*100) if n else 0

    # 8wk color
    w8_color = '#2563eb' if w8 and w8 >= 70 else ('#ea580c' if w8 and w8 >= 30 else '#dc2626')
    w52_color = '#2563eb' if w52 and w52 >= 50 else ('#ea580c' if w52 and w52 >= 20 else '#dc2626')
    w52_str = f'{w52}%' if w52 is not None else '—（資料不足）'
    nd_color = '#059669' if nd > 0 else '#dc2626'

    # Build replacement for this year's card stats div
    # We match the 4 stat divs inside the year's cohort card
    # Find by year label pattern
    year_label_pat = rf'{year}年加入（\d+人）'

    # Replace the 4 stats
    new_stats = (
        f'<div><span style="color:#666">第1週流失</span><br>'
        f'<strong style="font-size:1.3rem">{w1}人（{w1p}%）</strong></div>'
        f'<div><span style="color:#666">8週存活率</span><br>'
        f'<strong style="font-size:1.3rem;color:{w8_color}">{w8}%</strong></div>'
        f'<div><span style="color:#666">52週存活率</span><br>'
        f'<strong style="font-size:1.3rem;color:{w52_color}">{w52_str}</strong></div>'
        f'<div><span style="color:#666">從未流失</span><br>'
        f'<strong style="font-size:1.3rem;color:{nd_color}">{nd}人</strong></div>'
    )

    # Pattern: match the stats div inside this year's cohort card
    pat = (
        rf'({year}年加入（)\d+(人）</div>\s*'
        rf'<div style="display:flex[^>]*>)'
        rf'.*?'
        rf'(</div>\s*</div>\s*</div>)'
    )
    def replacer(m):
        return m.group(1) + str(n) + m.group(2) + new_stats + m.group(3)
    new_html = re.sub(pat, replacer, html, flags=re.DOTALL)
    if new_html == html:
        print(f"警告：無法替換 {year} cohort 卡")
    return new_html

for y in YEARS:
    html = update_cohort(html, y, year_stats[y])

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("index.html 已更新")
print(f"Summary: total={total}, w8={w8_pct}%, w1_dropout={w1_pct}%, active={active}")

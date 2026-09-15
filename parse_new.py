# -*- coding: utf-8 -*-
"""解析「更新后的课表大数据.xls」(2026-2027-1 大数据4班课表) → 生成三份数据
新导出格式：每页 11 行、3 周并排（每周 7 列），六大节各占一行。
输出：
  1) _new_weeks.json  —— index.html 内嵌 WEEKS 结构
  2) 秋季课表.json     —— 结构化 dict（term/class/school/source/timetable/weeks/courses）
  3) 秋季课表.md       —— 可读版
"""
import xlrd, json, re

SRC = r'C:\Users\30990\Desktop\头牟搭七\课表\更新后的课表大数据.xls'
DIR = r'C:\Users\30990\Desktop\头牟搭七\课表'

PT   = ['08:00-09:35', '09:50-11:25', '13:30-14:55',
        '15:10-16:35', '17:30-18:55', '19:10-20:35']
DAYS = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
PNAMES = ['第一大节', '第二大节', '第三大节', '第四大节', '第五大节', '第六大节']
PAGE_ROWS = 11

def parse_cell(raw):
    """一个单元格 → 若干条课程 entry。每个 entry 为 3~4 行：课程[类型]/老师/代码/教室(可空)。"""
    raw = raw.replace('\r\n', '\n').replace('\r', '\n')
    entries = []
    for block in re.split(r'\n\s*\n', raw):
        block = block.strip()
        if not block:
            continue
        lines = [l.strip() for l in block.split('\n') if l.strip()]
        if not lines:
            continue
        m = re.search(r'\[([^\]]*)\]', lines[0])
        ctype = m.group(1) if m else ''
        course = re.sub(r'\[[^\]]*\]', '', lines[0]).strip()
        teacher = lines[1] if len(lines) > 1 else ''
        code = lines[2] if len(lines) > 2 else ''
        room = lines[3] if len(lines) > 3 else ''
        room = room.replace('数据科学与大数据实验室', '').strip()
        entries.append({'course': course, 'type': ctype, 'teacher': teacher,
                        'code': code, 'room': room})
    return entries

wb = xlrd.open_workbook(SRC)
sh = wb.sheet_by_index(0)
n_pages = sh.nrows // PAGE_ROWS

weeks = []      # WEEKS 结构
courses = []    # 扁平 course 列表

for p in range(n_pages):
    title_row = p * PAGE_ROWS
    date_row = title_row + 2
    period_base = title_row + 4
    for wi in range(3):  # 页内 3 周
        wk = p * 3 + wi + 1
        dates = [str(sh.cell_value(date_row, wi * 7 + 1 + d)).strip() for d in range(7)]
        rng = f"{dates[0]} ~ {dates[6]}" if (dates[0] and dates[6]) else "待定"
        grid = []
        for di in range(7):  # 周一..周日
            day = []
            for pi in range(6):  # 六大节
                raw = str(sh.cell_value(period_base + pi, wi * 7 + 1 + di)).strip()
                if not raw:
                    day.append([])
                    continue
                es = parse_cell(raw)
                for e in es:
                    e['time'] = PT[pi]
                    courses.append({
                        'course': e['course'], 'type': e['type'], 'teacher': e['teacher'],
                        'code': e['code'], 'room': e['room'], 'time': e['time'],
                        'week': wk, 'date': dates[di], 'weekday': DAYS[di], 'period': PNAMES[pi],
                    })
                day.append([{'course': e['course'], 'teacher': e['teacher'],
                             'room': e['room'], 'time': e['time']} for e in es])
            grid.append(day)
        weeks.append({'week': wk, 'range': rng, 'dates': dates, 'grid': grid})

# 1) WEEKS
with open(DIR + r'\_new_weeks.json', 'w', encoding='utf-8') as f:
    json.dump(weeks, f, ensure_ascii=False)

# 2) 秋季课表.json
timetable = [
    {'period': '第一大节', 'section': '第1-2节',  'first': '08:00-09:35', 'other': '08:00-09:35'},
    {'period': '第二大节', 'section': '第3-4节',  'first': '10:00-11:35', 'other': '09:50-11:25'},
    {'period': '第三大节', 'section': '第5-6节',  'first': '13:30-14:55', 'other': '13:30-14:55'},
    {'period': '第四大节', 'section': '第7-8节',  'first': '15:10-16:35', 'other': '15:10-16:35'},
    {'period': '第五大节', 'section': '第9-10节', 'first': '17:30-18:55', 'other': '17:30-18:55'},
    {'period': '第六大节', 'section': '第11-12节','first': '19:10-20:35', 'other': '19:10-20:35'},
]
data = {
    'term': '2026-2027-1',
    'class': '24大数据4',
    'school': '哈尔滨石油学院',
    'source': '更新后的课表大数据.xls',
    'timetable': timetable,
    'weeks': [{'week': w['week'], 'dates': w['dates']} for w in weeks],
    'courses': courses,
}
with open(DIR + r'\秋季课表.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

# 3) 秋季课表.md
# 课程总览（按课程名聚合 类型）
from collections import OrderedDict
overview = OrderedDict()
for c in courses:
    key = c['course']
    if key not in overview:
        overview[key] = {'teacher': c['teacher'], 'types': []}
    if c['type'] and c['type'] not in overview[key]['types']:
        overview[key]['types'].append(c['type'])

lines = ['# 秋季课表 2026-2027-1（24大数据4）', '',
         '> 哈尔滨石油学院 · 数据来源 `更新后的课表大数据.xls` · 解析日期 2026-09-15', '',
         '## 课程总览', '', '| 课程 | 教师 | 类型 |', '|---|---|---|']
for name, info in overview.items():
    lines.append(f"| {name} | {info['teacher']} | {'/'.join(info['types']) or '-'} |")
lines.append('')

for w in weeks:
    # 收集该周有课的条目
    rows = []
    for c in courses:
        if c['week'] == w['week']:
            rows.append(f"- **{c['weekday']} {c['date']}** {c['period']} {c['time']} · {c['course']}（{c['teacher']}）`{c['room'] or '-'}`")
    lines.append(f"## 第{w['week']}周（{w['range'].replace(' ', '')}）")
    lines.append('')
    if rows:
        lines.extend(rows)
    else:
        lines.append('（无课）')
    lines.append('')

with open(DIR + r'\秋季课表.md', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print('weeks =', len(weeks), '| courses =', len(courses))
print('wrote _new_weeks.json / 秋季课表.json / 秋季课表.md')

# -*- coding: utf-8 -*-
"""解析室友课表 24电信2.xls → 生成 index.html 用的 WEEKS 结构（grid[day][period]）"""
import xlrd, json, re

SRC = r'C:\Users\30990\Desktop\24电信2.xls'
OUT = r'C:\Users\30990\Desktop\头牟搭七\课表\室友课表.json'

PT = ['08:00-09:35', '09:50-11:25', '13:30-14:55',
      '15:10-16:35', '17:30-18:55', '19:10-20:35']

wb = xlrd.open_workbook(SRC)
sh = wb.sheet_by_index(0)

n_weeks = sh.nrows // 10  # 每周 10 行
weeks = []
for wk in range(1, n_weeks + 1):
    base = (wk - 1) * 10
    # 标题(base+0) 学号(base+1) 表头(base+2) 日期(base+3) 六大节(base+4..9)
    dates = [str(sh.cell_value(base + 3, c)).strip() for c in range(1, 8)]
    rng = f"{dates[0]} ~ {dates[6]}" if (dates[0] and dates[6]) else "待定"

    grid = []  # 7 天
    for di in range(7):  # 周一..周日
        day = []
        for pi in range(6):  # 六大节
            raw = str(sh.cell_value(base + 4 + pi, di + 1)).strip()
            if not raw:
                day.append([])
                continue
            # 单元格：课程名[类型]\r\n老师\r\n代码\r\n教室
            parts = [p.strip() for p in raw.replace('\r\n', '\n').split('\n') if p.strip()]
            course = re.sub(r'\[[^\]]*\]', '', parts[0]).strip() if parts else ''
            teacher = parts[1] if len(parts) > 1 else ''
            room = parts[3] if len(parts) > 3 else ''
            day.append([{"course": course, "teacher": teacher, "room": room, "time": PT[pi]}])
        grid.append(day)
    weeks.append({"week": wk, "range": rng, "dates": dates, "grid": grid})

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(weeks, f, ensure_ascii=False, indent=2)

print('OK weeks=', len(weeks), '->', OUT)

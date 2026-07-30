# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

src = "D:/项目综合资料/天府办公文件/上海远香湖结算单/最总结算表格/02 结算前扣款事项汇总.xlsx"
out = "D:/27968/我的文档/Users/27968/Documents/青羊58施工计划项目/上海远香湖_物业备品配件数据说明.xlsx"

wb_src = openpyxl.load_workbook(src, data_only=True)
wb = openpyxl.Workbook()

hdr_font = Font(bold=True, size=10, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="4472C4")
red_fill = PatternFill("solid", fgColor="FFC7CE")
green_fill = PatternFill("solid", fgColor="C6EFCE")
yellow_fill = PatternFill("solid", fgColor="FFEB9C")
thin = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
title_font = Font(bold=True, size=14, color="003366")

def hdr(ws, row, cols):
    for c, h in enumerate(cols, 1):
        cell = ws.cell(row, c, h)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin

def sc(ws, row, cols):
    for c in range(1, cols+1):
        cell = ws.cell(row, c); cell.border = thin
        cell.alignment = Alignment(horizontal="center", vertical="center")

# ============ Sheet 1: 配品数据总说明 ============
ws1 = wb.active
ws1.title = "配品数据总说明"
ws1.cell(1, 1, "上海嘉定远香湖项目 — 物业配品配件数据说明").font = title_font
ws1.merge_cells("A1:I1")
ws1.cell(2, 1, "公式：配品(备件) + 合同清单量(合同量) = 应领总量(含损耗数量)  |  对比实际领用(四方单统计)").font = Font(size=10, italic=True, color="666666")
ws1.merge_cells("A2:I2")
ws1.cell(3, 1, "当 实际领用 < 应领总量 → 少领(需向甲方说明)；当 实际领用 > 应领总量 → 超领(需扣款)").font = Font(size=10, italic=True, color="FF0000")
ws1.merge_cells("A3:I3")

cols1 = ["材料类别", "材料名称", "规格型号", "单位", "合同清单量\n(合同量)", "物业配品配件\n(备件)", "应领总量=A+B\n(含损耗)", "实际领用\n(四方单统计)", "差额=C-D\n(正=少领/负=超领)", "少领/超领金额\n(含税)", "问题说明"]
hdr(ws1, 5, cols1)

row = 6
all_gap_items = []
total_less_qty = 0
total_less_amt = 0.0
total_over_qty = 0
total_over_amt = 0.0

def process_standard(ws_name, category, col_siji, col_hansun, col_price, start, end, price_lookup):
    global row, total_less_qty, total_less_amt, total_over_qty, total_over_amt
    ws_src = wb_src[ws_name]
    for r in range(start, end+1):
        name = ws_src.cell(r, 2).value
        spec = ws_src.cell(r, 3).value
        unit = ws_src.cell(r, 4).value
        shiji = ws_src.cell(r, col_siji).value
        hansun = ws_src.cell(r, col_hansun).value
        price = ws_src.cell(r, col_price).value
        if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
            hetong = hansun - shiji  # 配品 = 含损耗 - 四方单时的... wait
            # Actually: 含损耗 = 合同量 + 配品
            # 实际领用 = 四方单统计
            # 差额 = 含损耗 - 四方单统计
            # If 差额 > 0: 实际领用 < 应领总量 → 少领
            # If 差额 < 0: 实际领用 > 应领总量 → 超领
            chae = round(hansun - shiji, 2)
            
            # 配品 = 含损耗 - 合同? But I don't know合同... 
            # 从数据中，含损耗 = 合同量 + 配品
            # 合同量我不知道单独的值，但含损耗是已知的
            # 配品 = 含损耗 - 合同... 我不知道合同
            # Let me just use含损耗 as 应领总量
            peijian = None  # Unknown separately
            
            ws1.cell(row, 1, category)
            ws1.cell(row, 2, name)
            ws1.cell(row, 3, spec or "")
            ws1.cell(row, 4, unit or "")
            ws1.cell(row, 5, "")  # 合同量 unknown separately
            ws1.cell(row, 6, "")  # 配品 unknown separately
            ws1.cell(row, 7, hansun)
            ws1.cell(row, 8, shiji)
            ws1.cell(row, 9, chae)
            
            amount = 0
            if price and isinstance(price, (int,float)):
                amount = round(abs(chae) * price, 2)
            ws1.cell(row, 10, amount)
            
            if chae > 0:
                ws1.cell(row, 11, "少领: 实际领用少于应领总量，差额为剩余可移交物业的配件")
                total_less_qty += chae
                total_less_amt += amount
                for c in range(1, 12): ws1.cell(row, c).fill = green_fill
            elif chae < 0:
                ws1.cell(row, 11, "超领: 实际领用超过应领总量，超领部分需从结算中扣回")
                total_over_qty += abs(chae)
                total_over_amt += amount
                for c in range(1, 12): ws1.cell(row, c).fill = red_fill
            else:
                ws1.cell(row, 11, "持平: 实际领用等于应领总量")
            sc(ws1, row, 11)
            row += 1

# ===== 灯具 =====
ws_light = wb_src["灯具结算"]
for r in range(7, 36):
    name = ws_light.cell(r, 2).value
    spec = ws_light.cell(r, 3).value
    unit = ws_light.cell(r, 4).value
    shiji = ws_light.cell(r, 5).value
    hansun = ws_light.cell(r, 31).value
    price = ws_light.cell(r, 6).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        chae = round(hansun - shiji, 2)
        ws1.cell(row, 1, "灯具")
        ws1.cell(row, 2, name)
        ws1.cell(row, 3, spec or "")
        ws1.cell(row, 4, unit or "")
        ws1.cell(row, 5, "")  # 合同量不明
        ws1.cell(row, 6, "")  # 配品不明  
        ws1.cell(row, 7, hansun)
        ws1.cell(row, 8, shiji)
        ws1.cell(row, 9, chae)
        amount = round(abs(chae) * price, 2) if price and isinstance(price, (int,float)) else 0
        ws1.cell(row, 10, amount)
        if chae > 0:
            ws1.cell(row, 11, "少领")
            total_less_qty += chae; total_less_amt += amount
            for c in range(1,12): ws1.cell(row,c).fill = green_fill
        elif chae < 0:
            ws1.cell(row, 11, "超领")
            total_over_qty += abs(chae); total_over_amt += amount
            for c in range(1,12): ws1.cell(row,c).fill = red_fill
        else:
            ws1.cell(row, 11, "持平")
        sc(ws1, row, 11)
        row += 1

# ===== 开关插座面板 =====
ws_sw = wb_src["开关插座面板结算"]
for r in range(6, 39):
    name = ws_sw.cell(r, 2).value
    spec = ws_sw.cell(r, 3).value
    shiji = ws_sw.cell(r, 5).value
    hansun = ws_sw.cell(r, 30).value
    price = ws_sw.cell(r, 6).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        chae = round(hansun - shiji, 2)
        ws1.cell(row, 1, "开关插座面板")
        ws1.cell(row, 2, name)
        ws1.cell(row, 3, spec or "")
        ws1.cell(row, 4, "个")
        ws1.cell(row, 5, "")
        ws1.cell(row, 6, "")
        ws1.cell(row, 7, hansun)
        ws1.cell(row, 8, shiji)
        ws1.cell(row, 9, chae)
        amount = round(abs(chae) * price, 2) if price and isinstance(price, (int,float)) else 0
        ws1.cell(row, 10, amount)
        if chae > 0:
            ws1.cell(row, 11, "少领")
            total_less_qty += chae; total_less_amt += amount
            for c in range(1,12): ws1.cell(row,c).fill = green_fill
        elif chae < 0:
            ws1.cell(row, 11, "超领")
            total_over_qty += abs(chae); total_over_amt += amount
            for c in range(1,12): ws1.cell(row,c).fill = red_fill
        else:
            ws1.cell(row, 11, "持平")
        sc(ws1, row, 11)
        row += 1

# ===== 洁具科勒 =====
ws_k = wb_src["洁具结算（科勒）"]
for r in range(5, 10):
    name = ws_k.cell(r, 2).value
    shiji = ws_k.cell(r, 5).value
    heji = ws_k.cell(r, 13).value  # 合同合计 = 含损耗(此表无单独含损耗列)
    price = ws_k.cell(r, 6).value
    if name and isinstance(shiji, (int,float)) and isinstance(heji, (int,float)):
        chae = round(heji - shiji, 2)
        ws1.cell(row, 1, "洁具(科勒)")
        ws1.cell(row, 2, name)
        ws1.cell(row, 3, "")
        ws1.cell(row, 4, "套")
        ws1.cell(row, 5, "")
        ws1.cell(row, 6, "")
        ws1.cell(row, 7, heji)
        ws1.cell(row, 8, shiji)
        ws1.cell(row, 9, chae)
        amount = round(abs(chae) * price, 2) if price and isinstance(price, (int,float)) else 0
        ws1.cell(row, 10, amount)
        ws1.cell(row, 11, "持平" if chae == 0 else "少领" if chae > 0 else "超领")
        if chae > 0:
            total_less_qty += chae; total_less_amt += amount
        elif chae < 0:
            total_over_qty += abs(chae); total_over_amt += amount
        sc(ws1, row, 11)
        row += 1

# ===== HDPE管材结算 =====
ws_hdpe = wb_src["HDPE管材结算"]
for r in range(5, 51):
    name = ws_hdpe.cell(r, 3).value
    spec = ws_hdpe.cell(r, 4).value
    unit = ws_hdpe.cell(r, 6).value
    shiji = ws_hdpe.cell(r, 12).value  # 四方单
    hansun = ws_hdpe.cell(r, 20).value  # 含损耗数量
    price = ws_hdpe.cell(r, 9).value   # 含税单价
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        chae = round(hansun - shiji, 2)
        if chae != 0:  # 只显示有差异的
            ws1.cell(row, 1, "HDPE管材")
            ws1.cell(row, 2, name)
            ws1.cell(row, 3, spec or "")
            ws1.cell(row, 4, unit or "")
            ws1.cell(row, 5, "")
            ws1.cell(row, 6, "")
            ws1.cell(row, 7, hansun)
            ws1.cell(row, 8, shiji)
            ws1.cell(row, 9, chae)
            amount = round(abs(chae) * price, 2) if price and isinstance(price, (int,float)) else 0
            ws1.cell(row, 10, amount)
            if chae > 0:
                ws1.cell(row, 11, "少领"); total_less_qty += chae; total_less_amt += amount
                for c in range(1,12): ws1.cell(row,c).fill = green_fill
            else:
                ws1.cell(row, 11, "超领"); total_over_qty += abs(chae); total_over_amt += amount
                for c in range(1,12): ws1.cell(row,c).fill = red_fill
            sc(ws1, row, 11)
            row += 1

# ===== 汇总行 =====
sr = row + 1
ws1.cell(sr, 1, "=== 汇总 ===").font = Font(bold=True, size=12, color="003366")
ws1.merge_cells(f"A{sr}:D{sr}")
sr += 1
ws1.cell(sr, 1, "少领(实际<应领)汇总").font = Font(bold=True, size=11)
ws1.cell(sr, 7, int(total_less_qty)).font = Font(bold=True)
ws1.cell(sr, 10, round(total_less_amt, 2)).font = Font(bold=True, color="00FF0000")
for c in range(1,12): ws1.cell(sr,c).border = thin
sr += 1
ws1.cell(sr, 1, "超领(实际>应领)汇总").font = Font(bold=True, size=11)
ws1.cell(sr, 7, int(total_over_qty)).font = Font(bold=True)
ws1.cell(sr, 10, round(total_over_amt, 2)).font = Font(bold=True, color="00FF0000")
for c in range(1,12): ws1.cell(sr,c).border = thin

ws1.column_dimensions["A"].width = 16
ws1.column_dimensions["B"].width = 35
ws1.column_dimensions["C"].width = 25
ws1.column_dimensions["D"].width = 8
ws1.column_dimensions["E"].width = 16
ws1.column_dimensions["F"].width = 16
ws1.column_dimensions["G"].width = 16
ws1.column_dimensions["H"].width = 16
ws1.column_dimensions["I"].width = 12
ws1.column_dimensions["J"].width = 16
ws1.column_dimensions["K"].width = 50

# ============ Sheet 2: 甲供材超领扣款汇总 ============
ws2 = wb.create_sheet("超领-超供扣款汇总")
ws2.cell(1, 1, "甲供材超领(超供)扣款汇总").font = title_font
ws2.merge_cells("A1:E1")
cols2 = ["材料类别", "除税超领金额", "增值税", "含税超领金额", "备注"]
hdr(ws2, 3, cols2)

ws_sum = wb_src["甲供材超供汇总表"]
for r in range(4, 21):
    name = ws_sum.cell(r, 2).value
    hanhan = ws_sum.cell(r, 3).value
    rate = ws_sum.cell(r, 4).value
    hanhan_han = ws_sum.cell(r, 5).value
    note = ws_sum.cell(r, 9).value
    if name and isinstance(hanhan, (int,float)):
        ws2.cell(r, 1, name)
        ws2.cell(r, 2, hanhan)
        ws2.cell(r, 3, rate)
        ws2.cell(r, 4, hanhan_han)
        ws2.cell(r, 5, note or "")
        sc(ws2, r, 5)

ws2.column_dimensions["A"].width = 30
ws2.column_dimensions["B"].width = 18
ws2.column_dimensions["C"].width = 10
ws2.column_dimensions["D"].width = 18
ws2.column_dimensions["E"].width = 25

# ============ Sheet 3: 少领明细(可移交物业) ============
ws3 = wb.create_sheet("少领明细(可移交物业)")
ws3.cell(1, 1, "少领明细 — 实际领用少于应领总量(可移交物业的剩余配件)").font = title_font
ws3.merge_cells("A1:J1")
cols3 = ["材料类别", "材料名称", "单位", "应领总量(含损耗)", "实际领用(四方单)", "少领数量", "预估金额(含税)", "说明"]
hdr(ws3, 3, cols3)

r3 = 4
for r in range(6, ws1.max_row + 1):
    chae = ws1.cell(r, 9).value
    if isinstance(chae, (int,float)) and chae > 0:
        for c in range(1, 12):
            ws3.cell(r3, min(c, 8), ws1.cell(r, c).value)
        sc(ws3, r3, 8)
        r3 += 1

sr3 = r3 + 1
ws3.cell(sr3, 1, "合计").font = Font(bold=True, size=11)
ws3.cell(sr3, 4, int(total_less_qty)).font = Font(bold=True)
ws3.cell(sr3, 5, round(total_less_amt, 2)).font = Font(bold=True, color="00FF0000")
for c in range(1, 9): ws3.cell(sr3, c).border = thin

ws3.column_dimensions["A"].width = 16
ws3.column_dimensions["B"].width = 35
ws3.column_dimensions["C"].width = 8
ws3.column_dimensions["D"].width = 16
ws3.column_dimensions["E"].width = 16
ws3.column_dimensions["F"].width = 12
ws3.column_dimensions["G"].width = 16
ws3.column_dimensions["H"].width = 50

# ============ Sheet 4: 超领明细(需扣款) ============
ws4 = wb.create_sheet("超领明细(需扣款)")
ws4.cell(1, 1, "超领明细 — 实际领用超过应领总量(需从结算中扣回)").font = title_font
ws4.merge_cells("A1:J1")
cols4 = ["材料类别", "材料名称", "单位", "应领总量(含损耗)", "实际领用(四方单)", "超领数量", "预估金额(含税)", "说明"]
hdr(ws4, 3, cols4)

r4 = 4
for r in range(6, ws1.max_row + 1):
    chae = ws1.cell(r, 9).value
    if isinstance(chae, (int,float)) and chae < 0:
        for c in range(1, 12):
            ws4.cell(r4, min(c, 8), ws1.cell(r, c).value)
        sc(ws4, r4, 8)
        r4 += 1

wb.save(out)
print("文件已生成: " + out)
print(f"少领: {int(total_less_qty)}个/件, ¥{total_less_amt:.2f}")
print(f"超领: {int(total_over_qty)}个/件, ¥{total_over_amt:.2f}")

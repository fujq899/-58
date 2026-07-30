# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from copy import copy

out_path = r"D:\27968\我的文档\Users\27968\Documents\青羊58施工计划项目\上海远香湖_物业配件移交对比表.xlsx"
src_path = r"D:\项目综合资料\天府办公文件\上海远香湖结算单\最总结算表格\02 结算前扣款事项汇总.xlsx"

wb = openpyxl.load_workbook(out_path)
wb_src = openpyxl.load_workbook(src_path, data_only=True)

hdr_fill = PatternFill("solid", fgColor="4472C4")
hdr_font = Font(bold=True, size=11, color="FFFFFF")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
green_fill = PatternFill("solid", fgColor="C6EFCE")
red_fill = PatternFill("solid", fgColor="FFC7CE")

# 灯具单价表
light_prices = {}
for r in range(7, 36):
    name = wb_src["灯具结算"].cell(row=r, column=2).value
    price = wb_src["灯具结算"].cell(row=r, column=6).value
    if name and isinstance(price, (int,float)):
        light_prices[name] = price

# 开关面板单价表
switch_prices = {}
for r in range(6, 39):
    name = wb_src["开关插座面板结算"].cell(row=r, column=2).value
    price = wb_src["开关插座面板结算"].cell(row=r, column=6).value
    if name and isinstance(price, (int,float)):
        switch_prices[name] = price

# 洁具单价表
kohler_prices = {}
for r in range(5, 10):
    name = wb_src["洁具结算（科勒）"].cell(row=r, column=2).value
    price = wb_src["洁具结算（科勒）"].cell(row=r, column=6).value
    if name and isinstance(price, (int,float)):
        kohler_prices[name] = price

# 浴霸单价表
hb_prices = {}
for r in range(5, 8):
    name = wb_src["浴霸及凉霸结算"].cell(row=r, column=2).value
    price = wb_src["浴霸及凉霸结算"].cell(row=r, column=6).value
    if name and isinstance(price, (int,float)):
        hb_prices[name] = price

def get_price(cat, name):
    if cat == "灯具":
        return light_prices.get(name)
    elif cat == "开关插座面板":
        return switch_prices.get(name)
    elif "洁具" in cat:
        return kohler_prices.get(name)
    elif "浴霸" in cat:
        return hb_prices.get(name)
    return None

# ===== Sheet 1: 对比总表 - 添加预估金额列 =====
ws1 = wb["物业配品配件对比总表"]
# 在现有第8列后插入第9列"预估金额(含税)"
ws1.cell(3, 9, "预估金额(含税)").font = hdr_font
ws1.cell(3, 9).fill = hdr_fill
ws1.cell(3, 9).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws1.cell(3, 9).border = thin_border

for r in range(4, ws1.max_row + 1):
    cat = str(ws1.cell(r, 1).value or "")
    name = str(ws1.cell(r, 2).value or "")
    peijian = ws1.cell(r, 7).value or 0
    price = get_price(cat, name)
    if price and isinstance(peijian, (int,float)):
        if cat == "灯具" and "LED线条灯/公区" in name:
            price = 121.26  # 用LED线条灯正确的单价
        elif cat == "灯具" and "灯带驱动/户内及公区" in name:
            price = 74.70
            
        amount = abs(peijian) * price
        ws1.cell(r, 9, round(amount, 2))
    else:
        ws1.cell(r, 9, 0)
    ws1.cell(r, 9).border = thin_border
    ws1.cell(r, 9).alignment = Alignment(horizontal="center", vertical="center")
    if ws1.cell(r, 8).value and "超领" in str(ws1.cell(r, 8).value):
        ws1.cell(r, 9).fill = red_fill
    elif ws1.cell(r, 8).value and "有余量" in str(ws1.cell(r, 8).value):
        ws1.cell(r, 9).fill = green_fill

ws1.column_dimensions["I"].width = 16

# ===== Sheet 2: 对不上数量的明细 - 添加强调 =====
ws2 = wb["对不上数量的明细"]
# 同样加预估金额列
ws2.cell(3, 9, "预估金额(含税)").font = hdr_font
ws2.cell(3, 9).fill = hdr_fill
ws2.cell(3, 9).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws2.cell(3, 9).border = thin_border

for r in range(4, ws2.max_row + 1):
    cat = str(ws2.cell(r, 1).value or "")
    name = str(ws2.cell(r, 2).value or "")
    peijian = ws2.cell(r, 6).value or 0
    price = get_price(cat, name)
    if price and isinstance(peijian, (int,float)):
        amount = abs(peijian) * price
        ws2.cell(r, 9, round(amount, 2))
    else:
        ws2.cell(r, 9, 0)
    ws2.cell(r, 9).border = thin_border
    ws2.cell(r, 9).alignment = Alignment(horizontal="center", vertical="center")
    if ws2.cell(r, 7).value and "超领" in str(ws2.cell(r, 7).value):
        ws2.cell(r, 9).fill = red_fill
    elif ws2.cell(r, 7).value and "有余量" in str(ws2.cell(r, 7).value):
        ws2.cell(r, 9).fill = green_fill

ws2.column_dimensions["I"].width = 16

# ===== Sheet 3: 可移交物业配件清单 - 加金额列 =====
ws3 = wb["可移交物业配件清单"]
ws3.cell(3, 8, "预估金额(含税)").font = hdr_font
ws3.cell(3, 8).fill = hdr_fill
ws3.cell(3, 8).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws3.cell(3, 8).border = thin_border

total_qty = 0
total_amt = 0
for r in range(4, ws3.max_row + 1):
    cat = str(ws3.cell(r, 1).value or "")
    name = str(ws3.cell(r, 2).value or "")
    peijian = ws3.cell(r, 7).value or 0
    price = get_price(cat, name)
    if price and isinstance(peijian, (int,float)):
        amount = peijian * price
        ws3.cell(r, 8, round(amount, 2))
        total_qty += peijian
        total_amt += amount
    else:
        ws3.cell(r, 8, 0)
    ws3.cell(r, 8).border = thin_border
    ws3.cell(r, 8).alignment = Alignment(horizontal="center", vertical="center")

# 汇总行
sum_row = ws3.max_row + 2
ws3.cell(sum_row, 1, "合计").font = Font(bold=True, size=12, color="4472C4")
ws3.cell(sum_row, 1).border = thin_border
for c in range(2, 8):
    ws3.cell(sum_row, c).border = thin_border
ws3.cell(sum_row, 7, int(total_qty)).font = Font(bold=True, size=11)
ws3.cell(sum_row, 7).border = thin_border
ws3.cell(sum_row, 7).alignment = Alignment(horizontal="center")
ws3.cell(sum_row, 8, round(total_amt, 2)).font = Font(bold=True, size=11, color="red")
ws3.cell(sum_row, 8).border = thin_border
ws3.cell(sum_row, 8).alignment = Alignment(horizontal="center")

ws3.column_dimensions["H"].width = 16

wb.save(out_path)

# 输出汇总
print("=" * 55)
print("可移交物业配件 - 数量及金额汇总")
print("=" * 55)
print()
print("有余量可移交的共 %d项" % len(range(4, ws3.max_row + 1)))
print("可移交数量合计: %d个/件" % int(total_qty))
print("预估总金额(含税): ¥%.2f" % total_amt)
print()

# 输出超领项汇总
print("=" * 55)
print("超领(超出量) - 数量及金额汇总")
print("=" * 55)
print()
over_qty = 0
over_amt = 0
for r in range(4, ws1.max_row + 1):
    issue = ws1.cell(r, 8).value
    if issue and "超领" in str(issue):
        peijian = ws1.cell(r, 7).value or 0
        amount = ws1.cell(r, 9).value or 0
        over_qty += abs(peijian)
        over_amt += amount
print("超领项: 共%d项" % sum(1 for r in range(4, ws1.max_row + 1) if ws1.cell(r, 8).value and "超领" in str(ws1.cell(r, 8).value)))
print("超领数量合计: %d个/件" % int(over_qty))
print("超领预估金额(含税): ¥%.2f" % over_amt)
print()
print("文件已更新保存")

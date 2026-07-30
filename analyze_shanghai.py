# -*- coding: utf-8 -*-
import io, sys, openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

path = r'D:\项目综合资料\天府办公文件\上海远香湖结算单\最总结算表格\02 结算前扣款事项汇总.xlsx'
wb = openpyxl.load_workbook(path, data_only=True)

ws = wb['灯具结算']
print('【灯具】-- 超领明细（实际领用 > 含损耗合同量）：')
print('材料名称'.ljust(28), '单位'.ljust(6), '实际领用'.ljust(10), '合同含损耗'.ljust(12), '超领量'.ljust(10), '单价'.ljust(10), '超领金额'.ljust(12))
print('-' * 90)
for r in range(7, 36):
    name = ws.cell(row=r, column=2).value
    unit = ws.cell(row=r, column=4).value
    shiji = ws.cell(row=r, column=5).value
    hansun = ws.cell(row=r, column=31).value
    price = ws.cell(row=r, column=6).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        chaoling = shiji - hansun
        if chaoling > 0:
            amount = chaoling * price if isinstance(price, (int,float)) else 0
            pstr = '{:.2f}'.format(price) if isinstance(price, (int,float)) else '-'
            print('{:<28} {:<6} {:<10} {:<12.1f} {:<10.1f} {:<10} ¥{:.2f}'.format(str(name), str(unit) if unit else '', shiji, hansun, chaoling, pstr, amount))

print()
print('【灯具】-- 未超领项（实际领用 < 含损耗合同量，有余量）：')
print('材料名称'.ljust(28), '单位'.ljust(6), '实际领用'.ljust(10), '合同含损耗'.ljust(12), '余量'.ljust(10))
print('-' * 70)
for r in range(7, 36):
    name = ws.cell(row=r, column=2).value
    unit = ws.cell(row=r, column=4).value
    shiji = ws.cell(row=r, column=5).value
    hansun = ws.cell(row=r, column=31).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        yuliang = hansun - shiji
        if yuliang > 0:
            print('{:<28} {:<6} {:<10} {:<12.1f} {:<10.1f}'.format(str(name), str(unit) if unit else '', shiji, hansun, yuliang))

print()
ws3 = wb['开关插座面板结算']
print('【开关插座面板】-- 超领明细：')
print('材料名称'.ljust(28), '实际领用'.ljust(10), '合同含损耗'.ljust(12), '超领量'.ljust(10), '单价'.ljust(10), '超领金额'.ljust(12))
print('-' * 85)
for r in range(6, 39):
    name = ws3.cell(row=r, column=2).value
    shiji = ws3.cell(row=r, column=5).value
    hansun = ws3.cell(row=r, column=30).value
    price = ws3.cell(row=r, column=6).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        chaoling = shiji - hansun
        if chaoling > 0:
            amount = chaoling * price if isinstance(price, (int,float)) else 0
            pstr = '{:.2f}'.format(price) if isinstance(price, (int,float)) else '-'
            print('{:<28} {:<10} {:<12.1f} {:<10.1f} {:<10} ¥{:.2f}'.format(str(name), shiji, hansun, chaoling, pstr, amount))

print()
print('【开关插座面板】-- 未超领项（有余量）：')
print('材料名称'.ljust(28), '实际领用'.ljust(10), '合同含损耗'.ljust(12), '余量'.ljust(10))
print('-' * 65)
for r in range(6, 39):
    name = ws3.cell(row=r, column=2).value
    shiji = ws3.cell(row=r, column=5).value
    hansun = ws3.cell(row=r, column=30).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        yuliang = hansun - shiji
        if yuliang > 0:
            print('{:<28} {:<10} {:<12.1f} {:<10.1f}'.format(str(name), shiji, hansun, yuliang))

print()
print('【所有甲供材超领扣款汇总】')
print('材料类别'.ljust(28), '除税金额'.rjust(15), '含税金额'.rjust(15))
print('-' * 60)
ws_sum = wb['甲供材超供汇总表']
for r in range(4, 21):
    name = ws_sum.cell(row=r, column=2).value
    hanhan = ws_sum.cell(row=r, column=3).value
    hanhan_han = ws_sum.cell(row=r, column=5).value
    if name and isinstance(hanhan, (int,float)) and hanhan != 0:
        print('{:<28} ¥{:>12,.2f} ¥{:>12,.2f}'.format(str(name), hanhan, hanhan_han))

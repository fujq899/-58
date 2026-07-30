# -*- coding: utf-8 -*-
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import openpyxl

out_path = "D:/27968/我的文档/Users/27968/Documents/青羊58施工计划项目/上海远香湖_物业备品配件数据说明.xlsx"
src_path = "D:/项目综合资料/天府办公文件/上海远香湖结算单/最总结算表格/02 结算前扣款事项汇总.xlsx"

wb_src = openpyxl.load_workbook(src_path, data_only=True)

# 逐项读取所有数据
all_items = []
used_names = set()

# 灯具
ws = wb_src["灯具结算"]
for r in range(7, 36):
    name = ws.cell(r, 2).value
    spec = ws.cell(r, 3).value
    unit = ws.cell(r, 4).value
    shiji = ws.cell(r, 5).value
    price = ws.cell(r, 6).value
    hetong = ws.cell(r, 30).value  # 合同量
    hansun = ws.cell(r, 31).value  # 含损耗=合同+配品
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        ht = hetong if isinstance(hetong, (int,float)) else 0
        peijian = round(hansun - ht, 2) if isinstance(hetong, (int,float)) else 0
        chae = round(hansun - shiji, 2)
        yingling = hansun
        dj = price if isinstance(price, (int,float)) else 0
        all_items.append(("灯具", name, spec, unit or "", ht, peijian, yingling, shiji, chae, dj))
        used_names.add(name)

# 开关插座面板
ws = wb_src["开关插座面板结算"]
for r in range(6, 39):
    name = ws.cell(r, 2).value
    spec = ws.cell(r, 3).value
    shiji = ws.cell(r, 5).value
    price = ws.cell(r, 6).value
    hetong = ws.cell(r, 29).value
    hansun = ws.cell(r, 30).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        ht = hetong if isinstance(hetong, (int,float)) else 0
        peijian = round(hansun - ht, 2) if isinstance(hetong, (int,float)) else 0
        chae = round(hansun - shiji, 2)
        dj = price if isinstance(price, (int,float)) else 0
        all_items.append(("开关插座面板", name, spec or "", "个", ht, peijian, hansun, shiji, chae, dj))

# HDPE管材
ws = wb_src["HDPE管材结算"]
for r in range(5, 51):
    name = ws.cell(r, 3).value
    spec = str(ws.cell(r, 4).value or "")
    unit = str(ws.cell(r, 6).value or "")
    shiji = ws.cell(r, 12).value
    price = ws.cell(r, 9).value
    hetong = ws.cell(r, 15).value
    hansun = ws.cell(r, 20).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        ht = hetong if isinstance(hetong, (int,float)) else 0
        peijian = round(hansun - ht, 2) if isinstance(hetong, (int,float)) else 0
        chae = round(hansun - shiji, 2)
        dj = price if isinstance(price, (int,float)) else 0
        all_items.append(("HDPE管材", name, spec, unit, ht, peijian, hansun, shiji, chae, dj))

# 浴霸凉霸
ws = wb_src["浴霸及凉霸结算"]
for r in range(5, 8):
    name = ws.cell(r, 2).value
    shiji = ws.cell(r, 5).value
    price = ws.cell(r, 6).value
    heji = ws.cell(r, 13).value
    if name and isinstance(shiji, (int,float)) and isinstance(heji, (int,float)):
        chae = round(heji - shiji, 2)
        dj = price if isinstance(price, (int,float)) else 0
        all_items.append(("浴霸及凉霸", name, "", "只", heji, 0, heji, shiji, chae, dj))

# 洁具科勒
ws = wb_src["洁具结算（科勒）"]
for r in range(5, 10):
    name = ws.cell(r, 2).value
    shiji = ws.cell(r, 5).value
    price = ws.cell(r, 6).value
    heji = ws.cell(r, 13).value
    if name and isinstance(shiji, (int,float)) and isinstance(heji, (int,float)):
        chae = round(heji - shiji, 2)
        dj = price if isinstance(price, (int,float)) else 0
        all_items.append(("洁具(科勒)", name, "", "套", heji, 0, heji, shiji, chae, dj))

# 户内配电箱
ws = wb_src["户内配电箱结算"]
for r in range(5, 8):
    name = ws.cell(r, 2).value
    spec = ws.cell(r, 3).value
    shiji = ws.cell(r, 5).value
    price = ws.cell(r, 6).value
    heji = ws.cell(r, 13).value
    hansun = ws.cell(r, 14).value
    if name and isinstance(shiji, (int,float)) and isinstance(hansun, (int,float)):
        ht = heji if isinstance(heji, (int,float)) else 0
        chae = round(hansun - shiji, 2)
        dj = price if isinstance(price, (int,float)) else 0
        all_items.append(("户内配电箱", name, spec or "", "只", ht, round(hansun-ht,2), hansun, shiji, chae, dj))

# ===== 输出 =====
print("=" * 70)
print("上海嘉定远香湖项目 -- 物业配品配件移交数据说明")
print("=" * 70)
print()

# ------- 不够部分 (差额>0, 实际领用<应领) -------
print("【一、配品不够的项（实际领用 < 应领=合同+配品）】")
print("这些项的实际领用低于合同+配品的总量，意味着没有足够的配品可移交物业")
print()
print(f"  {'材料名称':<30} {'合同量':>8} {'+配品':>8} {'=应领':>8} {'-实际':>8} {'=不够':>8} {'单价':>8} {'金额':>10}")
print("  " + "-" * 90)

less_items = []
over_items = []

for item in sorted(all_items, key=lambda x: (x[0], x[1])):
    cat, name, spec, unit, ht, pj, yingling, shiji, chae, dj = item
    if chae > 0:
        amount = chae * dj
        less_items.append(item + (amount,))
        note_peijian = f"+{pj:.0f}" if pj > 0 else ""
        name_short = name[:30]
        print(f"  {name_short:<30} {ht:>8.0f} {note_peijian:>8} {yingling:>8.0f} {shiji:>8.0f} {chae:>8.0f} {dj:>8.2f} {amount:>10.2f}")

print()

less_gty = sum(i[8] for i in less_items)  # chae
less_amt = sum(i[10] for i in less_items)  # amount
print(f"  不够项小计: {len(less_items)}项, 不够数量={less_gty:.0f}个/件, 不够金额=¥{less_amt:.2f}")

print()
print("【二、配品足够的项（实际领用 >= 应领=合同+配品）】")
print("这些项的实际领用达到或超过了合同+配品的总量")
print()

for item in sorted(all_items, key=lambda x: (x[0], x[1])):
    cat, name, spec, unit, ht, pj, yingling, shiji, chae, dj = item
    if chae <= 0:
        amount = abs(chae) * dj
        over_items.append(item + (amount,))

over_gty = sum(abs(i[8]) for i in over_items)
over_amt = sum(i[10] for i in over_items)
print(f"  足够的项: {len(over_items)}项")
print(f"  其中超领(实际>应领): {sum(1 for i in over_items if i[8] < 0)}项, {over_gty:.0f}个/件, ¥{over_amt:.2f}(需扣款)")
print(f"  其中持平(实际=应领): {sum(1 for i in over_items if i[8] == 0)}项")

print()
print("=" * 70)
print("汇总")
print("=" * 70)
print(f"  A. 配品不够(实际领用<应领): {len(less_items)}项, 不够{less_gty:.0f}个, ¥{less_amt:.2f}")
print(f"  B. 配品足够(实际领用>=应领): {len(over_items)}项")
print(f"     - 其中超领的: {sum(1 for i in over_items if i[8] < 0)}项, 超领{over_gty:.0f}个, ¥{over_amt:.2f}")
print(f"     - 其中持平的: {sum(1 for i in over_items if i[8] == 0)}项")
print()
print(f"  ➜ 移交给物业的配品不够量: {less_gty:.0f}个/件, 价值¥{less_amt:.2f}")
print(f"  ➜ 超领(超用)量: {over_gty:.0f}个/件, 价值¥{over_amt:.2f}")

print()
print("=" * 70)
print("文件已保存: 上海远香湖_物业备品配件数据说明.xlsx")
print("  工作表1: 配品数据总说明 (所有材料完整对比)")
print("  工作表2: 少领明细(可移交物业) (不够的项)")
print("  工作表3: 超领明细(需扣款) (超用项)")

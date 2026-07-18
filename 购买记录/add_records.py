"""
购买记录归档脚本（完整版）
功能：
  1. 写入明细记录（序号用真实数字，不用公式）
  2. 自动重算并写入 月度统计/年度统计/平台分布 的数值
  3. 截图存入对应月份目录，"截图"列写入超链接
  4. 自动 git commit + push

用法:
  直接修改下方 RECORDS 列表，然后 python3 add_records.py
  或由 AI 在对话中调用本脚本的核心函数
"""
import os
import sys
import subprocess
from datetime import datetime
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList

XLSX = "/workspace/购买记录/归档表格/购买记录_2026.xlsx"
SHOT_BASE = "/workspace/购买记录/截图/2026"
YEAR = 2026

# ==================== 样式 ====================
COLOR_HEADER = "2F5496"
COLOR_TOTAL = "FFF2CC"
COLOR_BORDER = "8EAADB"
COLOR_TITLE = "1F4E79"
FONT_TITLE = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
FONT_HEADER = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
FONT_NORMAL = Font(name="微软雅黑", size=10, color="333333")
FONT_TOTAL = Font(name="微软雅黑", size=11, bold=True, color="C00000")
FONT_LINK = Font(name="微软雅黑", size=10, color="0563C1", underline="single")
FILL_HEADER = PatternFill("solid", fgColor=COLOR_HEADER)
FILL_TOTAL = PatternFill("solid", fgColor=COLOR_TOTAL)
FILL_TITLE = PatternFill("solid", fgColor=COLOR_TITLE)
thin_border = Border(
    left=Side(style="thin", color=COLOR_BORDER),
    right=Side(style="thin", color=COLOR_BORDER),
    top=Side(style="thin", color=COLOR_BORDER),
    bottom=Side(style="thin", color=COLOR_BORDER),
)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def get_existing_records(ws):
    """从明细表读取所有已有记录（用于重算统计）"""
    records = []
    row = 3
    while row <= 502:
        platform = ws.cell(row=row, column=2).value
        time_val = ws.cell(row=row, column=3).value
        if not time_val:
            break
        records.append({
            "platform": platform or "",
            "time": time_val,
            "product": ws.cell(row=row, column=4).value or "",
            "spec": ws.cell(row=row, column=5).value or "",
            "qty": ws.cell(row=row, column=6).value or 0,
            "paid": ws.cell(row=row, column=7).value or 0,
            "order": ws.cell(row=row, column=8).value or "",
        })
        row += 1
    return records


def unmerge_all(ws):
    """解除工作表所有合并单元格（重写前调用，避免 MergedCell 只读报错）"""
    merged = list(ws.merged_cells.ranges)
    for r in merged:
        ws.unmerge_cells(str(r))


def rewrite_detail_sheet(ws, records):
    """重写明细表：序号用真实数字"""
    unmerge_all(ws)
    # 清空数据区（第3行起）
    for row in range(3, 503):
        for col in range(1, 11):
            c = ws.cell(row=row, column=col)
            c.value = None
            c.hyperlink = None

    headers = ["序号", "购买平台", "下单时间", "商品名称", "规格", "数量", "实付金额(元)", "订单号", "截图", "备注"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = thin_border

    for i, rec in enumerate(records):
        row = 3 + i
        ws.cell(row=row, column=1, value=i + 1)  # 真实序号
        ws.cell(row=row, column=2, value=rec["platform"])
        ws.cell(row=row, column=3, value=rec["time"])
        ws.cell(row=row, column=3).number_format = "yyyy-mm-dd hh:mm:ss"
        ws.cell(row=row, column=4, value=rec["product"])
        ws.cell(row=row, column=5, value=rec["spec"])
        ws.cell(row=row, column=6, value=rec["qty"])
        ws.cell(row=row, column=7, value=rec["paid"])
        ws.cell(row=row, column=8, value=rec["order"])
        # 截图列：检查文件是否存在
        if rec.get("time"):
            t = rec["time"] if isinstance(rec["time"], datetime) else datetime.fromisoformat(str(rec["time"]))
            month_dir = os.path.join(SHOT_BASE, f"{t.month:02d}月")
            date_str = t.strftime("%Y%m%d")
            safe_product = rec["product"].replace("/", "_").replace("\\", "_")
            shot_filename = f"{rec['platform']}_{date_str}_{safe_product}.png"
            shot_full_path = os.path.join(month_dir, shot_filename)
            if os.path.exists(shot_full_path):
                cell = ws.cell(row=row, column=9)
                cell.value = "📸 查看"
                cell.hyperlink = f"file://{shot_full_path}"

        # 样式
        for col in range(1, 11):
            c = ws.cell(row=row, column=col)
            c.border = thin_border
            if col == 9:
                c.font = FONT_LINK
                c.alignment = ALIGN_CENTER
            else:
                c.font = FONT_NORMAL
                c.alignment = ALIGN_CENTER if col in (1, 2, 3, 6, 7) else ALIGN_LEFT
        ws.cell(row=row, column=7).number_format = "¥#,##0.00"
        ws.cell(row=row, column=6).number_format = "0"
        ws.row_dimensions[row].height = 22

    # 合计行
    total_row = 3 + len(records) + 1
    if len(records) > 0:
        ws.merge_cells(f"A{total_row}:F{total_row}")
        ws.cell(row=total_row, column=1, value="📊 合计").font = FONT_TOTAL
        ws.cell(row=total_row, column=1).fill = FILL_TOTAL
        ws.cell(row=total_row, column=1).alignment = ALIGN_CENTER
        total_paid = sum(r["paid"] for r in records)
        ws.cell(row=total_row, column=7, value=total_paid)
        ws.cell(row=total_row, column=7).font = FONT_TOTAL
        ws.cell(row=total_row, column=7).fill = FILL_TOTAL
        ws.cell(row=total_row, column=7).number_format = "¥#,##0.00"
        ws.cell(row=total_row, column=7).alignment = ALIGN_CENTER
        ws.cell(row=total_row, column=7).border = thin_border
        ws.merge_cells(f"H{total_row}:J{total_row}")
        ws.cell(row=total_row, column=8, value=f"{len(records)} 笔订单")
        ws.cell(row=total_row, column=8).font = FONT_TOTAL
        ws.cell(row=total_row, column=8).fill = FILL_TOTAL
        ws.cell(row=total_row, column=8).alignment = ALIGN_CENTER
        ws.cell(row=total_row, column=8).border = thin_border
        ws.row_dimensions[total_row].height = 28


def rewrite_month_sheet(ws, records):
    """重写月度统计：直接写入计算值"""
    unmerge_all(ws)
    # 清空
    for row in range(3, 20):
        for col in range(1, 7):
            ws.cell(row=row, column=col).value = None

    # 按月统计
    month_stats = defaultdict(lambda: {"count": 0, "amount": 0})
    for r in records:
        t = r["time"]
        if not t:
            continue
        t = t if isinstance(t, datetime) else datetime.fromisoformat(str(t))
        if t.year == YEAR:
            m = t.month
            month_stats[m]["count"] += 1
            month_stats[m]["amount"] += float(r["paid"] or 0)

    total_amount = sum(s["amount"] for s in month_stats.values())

    for m in range(1, 13):
        row = 2 + m
        s = month_stats[m]
        ws.cell(row=row, column=1, value=YEAR)
        ws.cell(row=row, column=2, value=f"{m:02d}月")
        ws.cell(row=row, column=3, value=s["count"])
        ws.cell(row=row, column=4, value=round(s["amount"], 2))
        ws.cell(row=row, column=5, value=round(s["amount"] / s["count"], 2) if s["count"] else 0)
        ws.cell(row=row, column=6, value=round(s["amount"] / total_amount, 4) if total_amount else 0)
        for col in range(1, 7):
            c = ws.cell(row=row, column=col)
            c.font = FONT_NORMAL
            c.border = thin_border
            c.alignment = ALIGN_CENTER
        ws.cell(row=row, column=4).number_format = "¥#,##0.00"
        ws.cell(row=row, column=5).number_format = "¥#,##0.00"
        ws.cell(row=row, column=6).number_format = "0.00%"

    # 合计
    total_row = 15
    total_count = sum(s["count"] for s in month_stats.values())
    ws.merge_cells(f"A{total_row}:B{total_row}")
    ws.cell(row=total_row, column=1, value="📊 合计")
    ws.cell(row=total_row, column=3, value=total_count)
    ws.cell(row=total_row, column=4, value=round(total_amount, 2))
    ws.cell(row=total_row, column=5, value=round(total_amount / total_count, 2) if total_count else 0)
    ws.cell(row=total_row, column=6, value=1.0)
    for col in range(1, 7):
        c = ws.cell(row=total_row, column=col)
        c.font = FONT_TOTAL
        c.fill = FILL_TOTAL
        c.border = thin_border
        c.alignment = ALIGN_CENTER
    ws.cell(row=total_row, column=4).number_format = "¥#,##0.00"
    ws.cell(row=total_row, column=5).number_format = "¥#,##0.00"
    ws.cell(row=total_row, column=6).number_format = "0.00%"

    # 清除旧图表，重新添加
    ws._charts = []

    # 月度金额柱状图
    bar = BarChart()
    bar.type = "col"
    bar.style = 10
    bar.title = f"{YEAR}年 月度购买金额趋势"
    bar.y_axis.title = "金额(元)"
    bar.x_axis.title = "月份"
    data = Reference(ws, min_col=4, min_row=2, max_row=14, max_col=4)
    cats = Reference(ws, min_col=2, min_row=3, max_row=14)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 10
    bar.width = 22
    ws.add_chart(bar, "H2")

    # 月度笔数折线图
    line = LineChart()
    line.title = f"{YEAR}年 月度购买笔数趋势"
    line.y_axis.title = "笔数"
    line.x_axis.title = "月份"
    data2 = Reference(ws, min_col=3, min_row=2, max_row=14, max_col=3)
    line.add_data(data2, titles_from_data=True)
    line.set_categories(cats)
    line.height = 10
    line.width = 22
    ws.add_chart(line, "H22")


def rewrite_year_sheet(ws, records):
    """重写年度统计"""
    unmerge_all(ws)
    for row in range(3, 10):
        for col in range(1, 6):
            ws.cell(row=row, column=col).value = None

    year_count = 0
    year_amount = 0
    for r in records:
        t = r["time"]
        if not t:
            continue
        t = t if isinstance(t, datetime) else datetime.fromisoformat(str(t))
        if t.year == YEAR:
            year_count += 1
            year_amount += float(r["paid"] or 0)

    row = 3
    ws.cell(row=row, column=1, value=YEAR)
    ws.cell(row=row, column=2, value=year_count)
    ws.cell(row=row, column=3, value=round(year_amount, 2))
    ws.cell(row=row, column=4, value=round(year_amount / year_count, 2) if year_count else 0)
    for col in range(1, 6):
        c = ws.cell(row=row, column=col)
        c.font = FONT_NORMAL
        c.border = thin_border
        c.alignment = ALIGN_CENTER
    ws.cell(row=row, column=3).number_format = "¥#,##0.00"
    ws.cell(row=row, column=4).number_format = "¥#,##0.00"

    # 柱状图
    ws._charts = []
    bar = BarChart()
    bar.type = "col"
    bar.style = 12
    bar.title = f"{YEAR}年 购买统计"
    bar.y_axis.title = "金额(元)"
    data = Reference(ws, min_col=3, min_row=2, max_row=3, max_col=3)
    cats = Reference(ws, min_col=1, min_row=3, max_row=3)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 8
    bar.width = 14
    bar.dataLabels = DataLabelList(showVal=True)
    ws.add_chart(bar, "G2")


def rewrite_platform_sheet(ws, records):
    """重写平台分布"""
    unmerge_all(ws)
    for row in range(3, 20):
        for col in range(1, 5):
            ws.cell(row=row, column=col).value = None

    platforms = ["抖音", "抖音商城", "1688", "淘宝", "京东", "拼多多", "天猫", "苏宁", "其他"]
    plat_stats = defaultdict(lambda: {"count": 0, "amount": 0})
    for r in records:
        t = r["time"]
        if not t:
            continue
        t = t if isinstance(t, datetime) else datetime.fromisoformat(str(t))
        if t.year != YEAR:
            continue
        p = r["platform"] if r["platform"] in platforms else "其他"
        plat_stats[p]["count"] += 1
        plat_stats[p]["amount"] += float(r["paid"] or 0)

    total_amount = sum(s["amount"] for s in plat_stats.values())
    plat_start = 3
    for i, p in enumerate(platforms):
        row = plat_start + i
        s = plat_stats[p]
        ws.cell(row=row, column=1, value=p)
        ws.cell(row=row, column=2, value=s["count"])
        ws.cell(row=row, column=3, value=round(s["amount"], 2))
        ws.cell(row=row, column=4, value=round(s["amount"] / total_amount, 4) if total_amount else 0)
        for col in range(1, 5):
            c = ws.cell(row=row, column=col)
            c.font = FONT_NORMAL
            c.border = thin_border
            c.alignment = ALIGN_CENTER
        ws.cell(row=row, column=3).number_format = "¥#,##0.00"
        ws.cell(row=row, column=4).number_format = "0.00%"

    plat_total_row = plat_start + len(platforms)
    total_count = sum(s["count"] for s in plat_stats.values())
    ws.cell(row=plat_total_row, column=1, value="📊 合计")
    ws.cell(row=plat_total_row, column=2, value=total_count)
    ws.cell(row=plat_total_row, column=3, value=round(total_amount, 2))
    ws.cell(row=plat_total_row, column=4, value=1.0 if total_amount else 0)
    for col in range(1, 5):
        c = ws.cell(row=plat_total_row, column=col)
        c.font = FONT_TOTAL
        c.fill = FILL_TOTAL
        c.border = thin_border
        c.alignment = ALIGN_CENTER
    ws.cell(row=plat_total_row, column=3).number_format = "¥#,##0.00"
    ws.cell(row=plat_total_row, column=4).number_format = "0.00%"

    # 图表
    ws._charts = []
    pie = PieChart()
    pie.title = f"{YEAR}年 各平台购买金额占比"
    pie_data = Reference(ws, min_col=3, min_row=2, max_row=plat_total_row - 1, max_col=3)
    pie_cats = Reference(ws, min_col=1, min_row=3, max_row=plat_total_row - 1)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.height = 10
    pie.width = 14
    pie.dataLabels = DataLabelList(showPercent=True)
    ws.add_chart(pie, "F2")

    bar = BarChart()
    bar.type = "bar"
    bar.style = 11
    bar.title = f"{YEAR}年 各平台购买笔数"
    bar.y_axis.title = "平台"
    bar.x_axis.title = "笔数"
    bdata = Reference(ws, min_col=2, min_row=2, max_row=plat_total_row - 1, max_col=2)
    bcats = Reference(ws, min_col=1, min_row=3, max_row=plat_total_row - 1)
    bar.add_data(bdata, titles_from_data=True)
    bar.set_categories(bcats)
    bar.height = 10
    bar.width = 14
    ws.add_chart(bar, "F22")


def add_records(new_records):
    """添加新记录并重算所有统计"""
    wb = load_workbook(XLSX)
    ws_detail = wb["🛒购买记录明细"]

    # 读取已有记录
    existing = get_existing_records(ws_detail)
    print(f"已有记录: {len(existing)} 条")

    # 追加新记录
    for rec in new_records:
        # 创建月份目录（按需）
        t = rec["time"]
        month_dir = os.path.join(SHOT_BASE, f"{t.month:02d}月")
        os.makedirs(month_dir, exist_ok=True)
        existing.append(rec)
        print(f"  + 新增: {rec['platform']} | {rec['time']} | {rec['product']} | ¥{rec['paid']}")

    # 重写所有工作表（用真实数值）
    rewrite_detail_sheet(ws_detail, existing)
    rewrite_month_sheet(wb["📊月度统计"], existing)
    rewrite_year_sheet(wb["📈年度统计"], existing)
    rewrite_platform_sheet(wb["🏪平台分布"], existing)

    wb.save(XLSX)
    print(f"\n✅ 已保存，共 {len(existing)} 条记录")


def git_commit_push():
    """git add + commit + push"""
    repo_dir = "/workspace"
    # add
    subprocess.run(["git", "add", "-A"], cwd=repo_dir, check=True)
    # 检查是否有变更
    result = subprocess.run(["git", "status", "--porcelain"], cwd=repo_dir, capture_output=True, text=True)
    if not result.stdout.strip():
        print("git: 无变更")
        return False
    # commit
    subprocess.run(
        ["git", "commit", "-m", f"feat: 更新购买记录 {datetime.now().strftime('%Y-%m-%d %H:%M')}"],
        cwd=repo_dir, check=True
    )
    # push
    push = subprocess.run(["git", "push"], cwd=repo_dir, capture_output=True, text=True)
    if push.returncode == 0:
        print("✅ git push 成功")
        return True
    else:
        print(f"❌ git push 失败: {push.stderr.strip()}")
        return False


# ==================== 待添加的记录 ====================
# AI 在对话中会修改这个列表，然后运行脚本
RECORDS = [
    {
        "platform": "1688",
        "time": datetime(2026, 7, 12, 12, 11, 38),
        "product": "创意世界杯足球大力神杯",
        "spec": "500ml透明圆形足球杯(牛皮纸盒包装)",
        "qty": 1,
        "paid": 5.90,
        "order": "3311879379799327755",
        "remark": "",
    },
    {
        "platform": "抖音商城",
        "time": datetime(2026, 7, 18, 10, 12, 44),
        "product": "创意大力神杯世界杯足球杯",
        "spec": "500ml琥珀圆形大力神杯(牛皮纸盒包装)",
        "qty": 1,
        "paid": 11.50,
        "order": "6928046994516508348",
        "remark": "",
    },
]


if __name__ == "__main__":
    # 检查是否已存在相同记录（避免重复）
    wb = load_workbook(XLSX)
    ws = wb["🛒购买记录明细"]
    existing = get_existing_records(ws)
    existing_orders = {r["order"] for r in existing}
    new_to_add = [r for r in RECORDS if r["order"] not in existing_orders]

    if new_to_add:
        print(f"新增 {len(new_to_add)} 条记录")
        add_records(new_to_add)
    else:
        print("记录已存在，仅重算统计")
        # 仍然重算，修复之前公式无值的问题
        add_records([])  # 传空列表，只重算

    # git
    git_commit_push()

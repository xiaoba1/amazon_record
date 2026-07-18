"""
购买记录 Excel 表格生成器（按年归档版）
用法:
    python3 generate_excel.py 2026     # 生成指定年度的表格
    python3 generate_excel.py          # 不带参数则生成当前年份

跨年时只需运行: python3 generate_excel.py 2027
即可在 归档表格/ 下生成新一年的空表格
"""
import os
import sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList

OUTPUT_DIR = "/workspace/购买记录/归档表格"

# ==================== 样式定义 ====================
COLOR_HEADER = "2F5496"
COLOR_HEADER_FONT = "FFFFFF"
COLOR_TOTAL = "FFF2CC"
COLOR_BORDER = "8EAADB"
COLOR_MONTH_ROW = "E2EFDA"
COLOR_TITLE = "1F4E79"

FONT_TITLE = Font(name="微软雅黑", size=16, bold=True, color="FFFFFF")
FONT_HEADER = Font(name="微软雅黑", size=11, bold=True, color=COLOR_HEADER_FONT)
FONT_NORMAL = Font(name="微软雅黑", size=10, color="333333")
FONT_TOTAL = Font(name="微软雅黑", size=11, bold=True, color="C00000")

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


def set_column_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def build_year_workbook(year: int):
    """生成指定年度的 Excel 表格"""
    wb = Workbook()
    DATA_START_ROW = 3
    DATA_END_ROW = 502  # 预留 500 行/年（日常足够，过多会拖慢重算）

    # ============================================================
    # Sheet 1: 🛒 购买记录明细
    # ============================================================
    ws1 = wb.active
    ws1.title = "🛒购买记录明细"

    ws1.merge_cells("A1:J1")
    ws1["A1"] = f"🛒 购买记录明细表  ({year}年)"
    ws1["A1"].font = FONT_TITLE
    ws1["A1"].fill = FILL_TITLE
    ws1["A1"].alignment = ALIGN_CENTER
    ws1.row_dimensions[1].height = 32

    headers = ["序号", "购买平台", "下单时间", "商品名称", "规格", "数量", "实付金额(元)", "订单号", "截图", "备注"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws1.cell(row=2, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = thin_border
    ws1.row_dimensions[2].height = 28

    set_column_widths(ws1, {
        "A": 8, "B": 12, "C": 22, "D": 32, "E": 22,
        "F": 8, "G": 14, "H": 24, "I": 12, "J": 20
    })

    # 截图超链接单元格的字体（蓝色下划线，模仿超链接样式）
    FONT_LINK = Font(name="微软雅黑", size=10, color="0563C1", underline="single")

    for row in range(DATA_START_ROW, DATA_END_ROW + 1):
        ws1.cell(row=row, column=1, value=f"=IF(C{row}=\"\",\"\",ROW()-2)")
        ws1.cell(row=row, column=7).number_format = "¥#,##0.00"
        ws1.cell(row=row, column=6).number_format = "0"
        for col in range(1, 11):
            c = ws1.cell(row=row, column=col)
            c.border = thin_border
            if col == 9:  # 截图列
                c.font = FONT_LINK
                c.alignment = ALIGN_CENTER
            else:
                c.font = FONT_NORMAL
                c.alignment = ALIGN_CENTER if col in (1, 2, 3, 6, 7) else ALIGN_LEFT
        ws1.row_dimensions[row].height = 22

    TOTAL_ROW = DATA_END_ROW + 2
    ws1.merge_cells(f"A{TOTAL_ROW}:F{TOTAL_ROW}")
    ws1.cell(row=TOTAL_ROW, column=1, value="📊 合计").font = FONT_TOTAL
    ws1.cell(row=TOTAL_ROW, column=1).fill = FILL_TOTAL
    ws1.cell(row=TOTAL_ROW, column=1).alignment = ALIGN_CENTER
    ws1.cell(row=TOTAL_ROW, column=7, value=f"=SUM(G{DATA_START_ROW}:G{DATA_END_ROW})")
    ws1.cell(row=TOTAL_ROW, column=7).font = FONT_TOTAL
    ws1.cell(row=TOTAL_ROW, column=7).fill = FILL_TOTAL
    ws1.cell(row=TOTAL_ROW, column=7).number_format = "¥#,##0.00"
    ws1.cell(row=TOTAL_ROW, column=7).alignment = ALIGN_CENTER
    ws1.cell(row=TOTAL_ROW, column=7).border = thin_border
    ws1.merge_cells(f"H{TOTAL_ROW}:J{TOTAL_ROW}")
    ws1.cell(row=TOTAL_ROW, column=8, value=f"=COUNTA(C{DATA_START_ROW}:C{DATA_END_ROW}) & \" 笔订单\"")
    ws1.cell(row=TOTAL_ROW, column=8).font = FONT_TOTAL
    ws1.cell(row=TOTAL_ROW, column=8).fill = FILL_TOTAL
    ws1.cell(row=TOTAL_ROW, column=8).alignment = ALIGN_CENTER
    ws1.cell(row=TOTAL_ROW, column=8).border = thin_border
    ws1.row_dimensions[TOTAL_ROW].height = 28
    ws1.freeze_panes = "A3"

    # ============================================================
    # Sheet 2: 📊 月度统计（仅当年 12 个月）
    # ============================================================
    ws2 = wb.create_sheet("📊月度统计")
    ws2.merge_cells("A1:F1")
    ws2["A1"] = f"📊 月度购买统计 ({year}年)"
    ws2["A1"].font = FONT_TITLE
    ws2["A1"].fill = FILL_TITLE
    ws2["A1"].alignment = ALIGN_CENTER
    ws2.row_dimensions[1].height = 32

    month_headers = ["年份", "月份", "购买笔数", "总金额(元)", "平均单价(元)", "占比"]
    for col_idx, h in enumerate(month_headers, 1):
        cell = ws2.cell(row=2, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = thin_border
    ws2.row_dimensions[2].height = 28
    set_column_widths(ws2, {"A": 10, "B": 10, "C": 12, "D": 16, "E": 16, "F": 10})

    start_row = 3
    for m in range(1, 13):
        row = start_row + m - 1
        ws2.cell(row=row, column=1, value=year)
        ws2.cell(row=row, column=2, value=f"{m:02d}月")
        # 用 COUNTIFS/SUMIFS 按日期范围统计（比 SUMPRODUCT 快得多）
        m_start = f"DATE({year},{m},1)"
        m_end = f"DATE({year},{m+1 if m<12 else 1},1)" + (f"+365" if m == 12 else "")
        ws2.cell(row=row, column=3, value=(
            f'=COUNTIFS(\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},">="&{m_start},'
            f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},"<"&{m_end})'
        ))
        ws2.cell(row=row, column=4, value=(
            f'=SUMIFS(\'🛒购买记录明细\'!G{DATA_START_ROW}:G{DATA_END_ROW},'
            f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},">="&{m_start},'
            f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},"<"&{m_end})'
        ))
        ws2.cell(row=row, column=5, value=f'=IF(C{row}=0,0,D{row}/C{row})')
        ws2.cell(row=row, column=6, value=f'=IF(D{row}=0,0,D{row}/SUM($D${start_row}:$D${start_row+11}))')
        for col in range(1, 7):
            c = ws2.cell(row=row, column=col)
            c.font = FONT_NORMAL
            c.border = thin_border
            c.alignment = ALIGN_CENTER
        ws2.cell(row=row, column=4).number_format = "¥#,##0.00"
        ws2.cell(row=row, column=5).number_format = "¥#,##0.00"
        ws2.cell(row=row, column=6).number_format = "0.00%"
        ws2.row_dimensions[row].height = 22

    month_total_row = start_row + 12
    ws2.merge_cells(f"A{month_total_row}:B{month_total_row}")
    ws2.cell(row=month_total_row, column=1, value="📊 合计").font = FONT_TOTAL
    ws2.cell(row=month_total_row, column=1).fill = FILL_TOTAL
    ws2.cell(row=month_total_row, column=1).alignment = ALIGN_CENTER
    ws2.cell(row=month_total_row, column=3, value=f"=SUM(C{start_row}:C{month_total_row-1})")
    ws2.cell(row=month_total_row, column=4, value=f"=SUM(D{start_row}:D{month_total_row-1})")
    ws2.cell(row=month_total_row, column=5, value=f"=IF(C{month_total_row}=0,0,D{month_total_row}/C{month_total_row})")
    ws2.cell(row=month_total_row, column=6, value="100.00%")
    for col in range(1, 7):
        c = ws2.cell(row=month_total_row, column=col)
        c.font = FONT_TOTAL
        c.fill = FILL_TOTAL
        c.border = thin_border
        c.alignment = ALIGN_CENTER
    ws2.cell(row=month_total_row, column=4).number_format = "¥#,##0.00"
    ws2.cell(row=month_total_row, column=5).number_format = "¥#,##0.00"
    ws2.row_dimensions[month_total_row].height = 28

    # 月度金额柱状图
    bar_chart = BarChart()
    bar_chart.type = "col"
    bar_chart.style = 10
    bar_chart.title = f"{year}年 月度购买金额趋势"
    bar_chart.y_axis.title = "金额(元)"
    bar_chart.x_axis.title = "月份"
    data_ref = Reference(ws2, min_col=4, min_row=2, max_row=month_total_row-1, max_col=4)
    cats_ref = Reference(ws2, min_col=2, min_row=3, max_row=month_total_row-1)
    bar_chart.add_data(data_ref, titles_from_data=True)
    bar_chart.set_categories(cats_ref)
    bar_chart.height = 10
    bar_chart.width = 22
    ws2.add_chart(bar_chart, "H2")

    line_chart = LineChart()
    line_chart.title = f"{year}年 月度购买笔数趋势"
    line_chart.y_axis.title = "笔数"
    line_chart.x_axis.title = "月份"
    data_ref2 = Reference(ws2, min_col=3, min_row=2, max_row=month_total_row-1, max_col=3)
    line_chart.add_data(data_ref2, titles_from_data=True)
    line_chart.set_categories(cats_ref)
    line_chart.height = 10
    line_chart.width = 22
    ws2.add_chart(line_chart, "H22")
    ws2.freeze_panes = "A3"

    # ============================================================
    # Sheet 3: 📈 年度统计（当年单行 + 合计）
    # ============================================================
    ws3 = wb.create_sheet("📈年度统计")
    ws3.merge_cells("A1:E1")
    ws3["A1"] = f"📈 年度购买统计 ({year}年)"
    ws3["A1"].font = FONT_TITLE
    ws3["A1"].fill = FILL_TITLE
    ws3["A1"].alignment = ALIGN_CENTER
    ws3.row_dimensions[1].height = 32

    year_headers = ["年份", "购买笔数", "总金额(元)", "平均单价(元)", "备注"]
    for col_idx, h in enumerate(year_headers, 1):
        cell = ws3.cell(row=2, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = thin_border
    ws3.row_dimensions[2].height = 28
    set_column_widths(ws3, {"A": 10, "B": 12, "C": 16, "D": 16, "E": 20})

    row = 3
    ws3.cell(row=row, column=1, value=year)
    ws3.cell(row=row, column=2, value=(
        f'=COUNTIFS(\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},">="&DATE({year},1,1),'
        f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},"<"&DATE({year+1},1,1))'
    ))
    ws3.cell(row=row, column=3, value=(
        f'=SUMIFS(\'🛒购买记录明细\'!G{DATA_START_ROW}:G{DATA_END_ROW},'
        f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},">="&DATE({year},1,1),'
        f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},"<"&DATE({year+1},1,1))'
    ))
    ws3.cell(row=row, column=4, value=f'=IF(B{row}=0,0,C{row}/B{row})')
    for col in range(1, 6):
        c = ws3.cell(row=row, column=col)
        c.font = FONT_NORMAL
        c.border = thin_border
        c.alignment = ALIGN_CENTER
    ws3.cell(row=row, column=3).number_format = "¥#,##0.00"
    ws3.cell(row=row, column=4).number_format = "¥#,##0.00"
    ws3.row_dimensions[row].height = 24
    ws3.freeze_panes = "A3"

    # ============================================================
    # Sheet 4: 🏪 平台分布
    # ============================================================
    ws4 = wb.create_sheet("🏪平台分布")
    ws4.merge_cells("A1:D1")
    ws4["A1"] = f"🏪 购买平台分布 ({year}年)"
    ws4["A1"].font = FONT_TITLE
    ws4["A1"].fill = FILL_TITLE
    ws4["A1"].alignment = ALIGN_CENTER
    ws4.row_dimensions[1].height = 32

    platform_headers = ["购买平台", "购买笔数", "总金额(元)", "占比"]
    for col_idx, h in enumerate(platform_headers, 1):
        cell = ws4.cell(row=2, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = thin_border
    ws4.row_dimensions[2].height = 28
    set_column_widths(ws4, {"A": 14, "B": 12, "C": 16, "D": 10})

    platforms = ["抖音", "1688", "淘宝", "京东", "拼多多", "天猫", "苏宁", "其他"]
    plat_start = 3
    # 平台统计：用 COUNTIFS/SUMIFS（平台 + 当年日期范围），比 SUMPRODUCT 快得多
    for i, p in enumerate(platforms):
        row = plat_start + i
        ws4.cell(row=row, column=1, value=p)
        ws4.cell(row=row, column=2, value=(
            f'=COUNTIFS(\'🛒购买记录明细\'!B{DATA_START_ROW}:B{DATA_END_ROW},"{p}",'
            f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},">="&DATE({year},1,1),'
            f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},"<"&DATE({year+1},1,1))'
        ))
        ws4.cell(row=row, column=3, value=(
            f'=SUMIFS(\'🛒购买记录明细\'!G{DATA_START_ROW}:G{DATA_END_ROW},'
            f'\'🛒购买记录明细\'!B{DATA_START_ROW}:B{DATA_END_ROW},"{p}",'
            f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},">="&DATE({year},1,1),'
            f'\'🛒购买记录明细\'!C{DATA_START_ROW}:C{DATA_END_ROW},"<"&DATE({year+1},1,1))'
        ))
        ws4.cell(row=row, column=4, value=f'=IF(C{row}=0,0,C{row}/SUM($C${plat_start}:$C${plat_start+len(platforms)-1}))')
        for col in range(1, 5):
            c = ws4.cell(row=row, column=col)
            c.font = FONT_NORMAL
            c.border = thin_border
            c.alignment = ALIGN_CENTER
        ws4.cell(row=row, column=3).number_format = "¥#,##0.00"
        ws4.cell(row=row, column=4).number_format = "0.00%"
        ws4.row_dimensions[row].height = 22

    plat_total_row = plat_start + len(platforms)
    ws4.cell(row=plat_total_row, column=1, value="📊 合计").font = FONT_TOTAL
    ws4.cell(row=plat_total_row, column=1).fill = FILL_TOTAL
    ws4.cell(row=plat_total_row, column=2, value=f"=SUM(B{plat_start}:B{plat_total_row-1})")
    ws4.cell(row=plat_total_row, column=3, value=f"=SUM(C{plat_start}:C{plat_total_row-1})")
    ws4.cell(row=plat_total_row, column=4, value="100.00%")
    for col in range(1, 5):
        c = ws4.cell(row=plat_total_row, column=col)
        c.font = FONT_TOTAL
        c.fill = FILL_TOTAL
        c.border = thin_border
        c.alignment = ALIGN_CENTER
    ws4.cell(row=plat_total_row, column=3).number_format = "¥#,##0.00"
    ws4.row_dimensions[plat_total_row].height = 28

    # 平台金额饼图
    pie = PieChart()
    pie.title = f"{year}年 各平台购买金额占比"
    pie_data = Reference(ws4, min_col=3, min_row=2, max_row=plat_total_row-1, max_col=3)
    pie_cats = Reference(ws4, min_col=1, min_row=3, max_row=plat_total_row-1)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.height = 10
    pie.width = 14
    pie.dataLabels = DataLabelList(showPercent=True)
    ws4.add_chart(pie, "F2")

    plat_bar = BarChart()
    plat_bar.type = "bar"
    plat_bar.style = 11
    plat_bar.title = f"{year}年 各平台购买笔数"
    plat_bar.y_axis.title = "平台"
    plat_bar.x_axis.title = "笔数"
    pb_data = Reference(ws4, min_col=2, min_row=2, max_row=plat_total_row-1, max_col=2)
    pb_cats = Reference(ws4, min_col=1, min_row=3, max_row=plat_total_row-1)
    plat_bar.add_data(pb_data, titles_from_data=True)
    plat_bar.set_categories(pb_cats)
    plat_bar.height = 10
    plat_bar.width = 14
    ws4.add_chart(plat_bar, "F22")
    ws4.freeze_panes = "A3"

    return wb


def main():
    # 解析年份参数
    if len(sys.argv) >= 2:
        year = int(sys.argv[1])
    else:
        year = datetime.now().year

    output_file = os.path.join(OUTPUT_DIR, f"购买记录_{year}.xlsx")

    if os.path.exists(output_file):
        print(f"⚠️  文件已存在: {output_file}")
        print(f"    如需重新生成，请先删除或备份该文件。")
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    # 仅创建年度截图目录，月份子目录在录入截图时按需创建
    shot_year_dir = f"/workspace/购买记录/截图/{year}"
    os.makedirs(shot_year_dir, exist_ok=True)

    wb = build_year_workbook(year)
    wb.save(output_file)

    print(f"✅ 已生成 {year} 年度表格: {output_file}")
    print(f"   包含 4 个工作表:")
    print(f"   1. 🛒购买记录明细 - 详细记录(预留500行，含截图超链接列)")
    print(f"   2. 📊月度统计 - 含柱状图+折线图")
    print(f"   3. 📈年度统计 - 当年总览")
    print(f"   4. 🏪平台分布 - 含饼图+柱状图")
    print(f"   截图目录: {shot_year_dir}/ (月份子目录按需创建)")


if __name__ == "__main__":
    main()

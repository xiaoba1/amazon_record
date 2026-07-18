"""
将已确认的购买记录写入 2026 年度 Excel 明细表
并保存截图到对应月份目录、在"截图"列写入超链接
"""
import os
import shutil
from datetime import datetime
from openpyxl import load_workbook

XLSX = "/workspace/购买记录/归档表格/购买记录_2026.xlsx"
SHOT_BASE = "/workspace/购买记录/截图/2026"

# 已确认的两条记录
# screenshot_path: 原图路径（由用户上传，此处用占位；实际写入时需替换为真实路径）
# 如果用户没提供真实图片文件，则只写记录，截图列留空
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

wb = load_workbook(XLSX)
ws = wb["🛒购买记录明细"]

# 找第一个空行（C列下单时间为空即为空行）
DATA_START = 3
DATA_END = 502
row = DATA_START
while row <= DATA_END:
    if ws.cell(row=row, column=3).value in (None, ""):
        break
    row += 1

if row > DATA_END:
    raise RuntimeError("明细表已满，请扩大预留行数")

for rec in RECORDS:
    month_dir = os.path.join(SHOT_BASE, f"{rec['time'].month:02d}月")
    os.makedirs(month_dir, exist_ok=True)

    # 截图文件名：平台_日期_商品.png
    date_str = rec["time"].strftime("%Y%m%d")
    # 商品名做简单清洗，避免文件名非法字符
    safe_product = rec["product"].replace("/", "_").replace("\\", "_")
    shot_filename = f"{rec['platform']}_{date_str}_{safe_product}.png"
    shot_full_path = os.path.join(month_dir, shot_filename)

    # 写入数据
    # A=序号(公式), B=平台, C=下单时间, D=商品名称, E=规格, F=数量, G=实付, H=订单号, I=截图, J=备注
    ws.cell(row=row, column=2, value=rec["platform"])
    ws.cell(row=row, column=3, value=rec["time"])
    ws.cell(row=row, column=3).number_format = "yyyy-mm-dd hh:mm:ss"
    ws.cell(row=row, column=4, value=rec["product"])
    ws.cell(row=row, column=5, value=rec["spec"])
    ws.cell(row=row, column=6, value=rec["qty"])
    ws.cell(row=row, column=7, value=rec["paid"])
    ws.cell(row=row, column=8, value=rec["order"])

    # 截图超链接（仅当截图文件真实存在时才设置超链接）
    if os.path.exists(shot_full_path):
        cell = ws.cell(row=row, column=9)
        cell.value = "📸 查看"
        cell.hyperlink = f"file://{shot_full_path}"
    # 否则截图列留空（暂无图片文件）

    if rec["remark"]:
        ws.cell(row=row, column=10, value=rec["remark"])

    print(f"✅ 第{row-2}条: {rec['platform']} | {rec['time']} | {rec['product']} | ¥{rec['paid']}")
    print(f"   截图将存放: {shot_full_path}" + (" (图片已就位)" if os.path.exists(shot_full_path) else " (暂无图片文件)"))
    row += 1

wb.save(XLSX)
print(f"\n已保存 {len(RECORDS)} 条记录到: {XLSX}")

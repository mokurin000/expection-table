import json
import math
from typing import Iterable

import xlsxwriter

# 加载 JSON 数据
with open("plants.json", "r", encoding="utf-8") as f:
    plants_data = json.load(f)


# 分离地球和月球作物
def sort_plants(plants: Iterable[dict]) -> list:
    return sorted(
        plants,
        key=lambda p: p["maxWeight"] ** 1.5 * p["priceCoefficient"],
        reverse=True,
    )


earth_plants = sort_plants(p for p in plants_data if p["type"] == "普通")
moon_plants = sort_plants(p for p in plants_data if p["type"] == "月球")


# 基础突变类型及倍数
mutations_earth = [
    {"name": "无", "mult": 1.0},
    {"name": "银", "mult": 3.0},
    {"name": "金", "mult": 10.0},
    {"name": "水晶", "mult": 20.0},
    {"name": "流光", "mult": 30.0},
]

mutations_moon = mutations_earth + [{"name": "星空", "mult": 40.0}]

# 洒水器类型及对应的 k 值（空刷=1, 简易=1.2, 标准=1.7, 白银=2.4, 黄金=3.4）
sprinklers = [
    {"name": "空刷", "k": 1.0},
    {"name": "简易", "k": 1.2},
    {"name": "标准", "k": 1.7},
    {"name": "白银", "k": 2.4},
    {"name": "黄金", "k": 3.4},
]


def format_price(value):
    """格式化价格：≥10000 显示为 X.X万（保留3位小数），否则完整数字"""
    if value >= 10000:
        v = value / 10000
        return f"{v:.2f}万"
    else:
        return f"{int(round(value)):,}"


def create_sheet(workbook, sheet_name, plants, mutations):
    worksheet = workbook.add_worksheet(sheet_name)

    headers = ["作物名称", "突变", "规模", "空刷", "简易", "标准", "白银", "黄金"]

    # 洒水器列背景色
    colors = ["#FFFFFF", "#D3D3D3", "#90EE90", "#87CEFA", "#FFCC66"]

    # 普通行格式（细框）
    col_formats = [
        workbook.add_format(
            {
                "bg_color": c,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
            }
        )
        for c in colors
    ]

    # 作物首行格式（粗上框）
    col_formats_top = [
        workbook.add_format(
            {
                "bg_color": c,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
                "top": 2,
            }
        )
        for c in colors
    ]

    header_fmt = workbook.add_format(
        {"bold": True, "align": "center", "valign": "vcenter", "border": 1}
    )

    cell_fmt = workbook.add_format({"border": 1})
    cell_fmt_top = workbook.add_format({"border": 1, "top": 2})

    # 写入表头
    for col, header in enumerate(headers):
        if col < 3:
            worksheet.write(0, col, header, header_fmt)
        else:
            h_fmt = workbook.add_format(
                {
                    "bg_color": colors[col - 3],
                    "bold": True,
                    "align": "center",
                    "valign": "vcenter",
                    "border": 1,
                }
            )
            worksheet.write(0, col, header, h_fmt)

    # 列宽
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 10)
    worksheet.set_column(2, 2, 10)
    worksheet.set_column(3, 7, 28)

    row = 1

    for plant in plants:
        name = plant["name"]
        max_weight = plant["maxWeight"]
        price_coeff = plant["priceCoefficient"]

        first_row_of_plant = True

        for mut in mutations:
            mult = mut["mult"]
            mut_name = mut["name"]

            for scale, is_giant in [("普通", False), ("巨大", True)]:
                fmt = cell_fmt_top if first_row_of_plant else cell_fmt

                worksheet.write(row, 0, name, fmt)
                worksheet.write(row, 1, mut_name, fmt)
                worksheet.write(row, 2, scale, fmt)

                G = 5.0 if is_giant else 1.0

                for s_idx, sp in enumerate(sprinklers):
                    k = sp["k"]

                    effective_w = max_weight * k * G / 34
                    const = (2 / 5) * (4 * math.sqrt(2) - 1)
                    expected = const * price_coeff * (effective_w**1.5) * mult

                    cell_text = format_price(expected)

                    fmt_list = col_formats_top if first_row_of_plant else col_formats
                    worksheet.write(row, s_idx + 3, cell_text, fmt_list[s_idx])

                first_row_of_plant = False
                row += 1

    # 冻结前三列
    worksheet.freeze_panes(1, 3)


# 生成 Excel
workbook = xlsxwriter.Workbook("作物期望价值表_分行版.xlsx")

create_sheet(workbook, "地球", earth_plants, mutations_earth)
create_sheet(workbook, "月球", moon_plants, mutations_moon)

workbook.close()

print("✅ 已生成：作物期望价值表_分行版.xlsx")
print("   • 每个作物第一行：粗上框线")
print("   • 其余所有单元格：细框线")
print("   • 普通行 G=1，巨大行 G=5")
print("   • 每格为期望价值（已乘突变倍数）")

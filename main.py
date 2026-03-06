import json
import math
from typing import Iterable
import xlsxwriter

# 加载 JSON 数据
with open("plants.json", "r", encoding="utf-8") as f:
    plants_data = json.load(f)

# 稀有度分组（请确保与你的 plants.json 中的 name 完全匹配）
# fmt: off
common = ["月光草", "灰壤豆", "土豆", "香菇", "番茄", "波斯菊"]
uncommon = ["月灯草", "月番茄", "大豆", "竹子", "黄瓜"]
rare = ["西瓜", "梨", "橘子", "玉米", "白菜", "牵牛花", "棉花", "月环树", "银灰苔"]
legendary = ["月莓", "星叶菜", "苹果", "石榴", "香蕉", "车厘子", "椰子", "南瓜"]
devine = ["草莓", "猕猴桃", "荔枝", "榴莲", "月核树", "液光藤"]
prismatic = [
    "月影梅", "幻月花", "星空玫瑰", "红包树", "月兔",
    "向日葵", "松果", "大王菊", "葡萄", "蟠桃",
    "惊奇菇", "仙人掌象", "魔鬼朝天椒"
]
# fmt: on


# 分离地球和月球作物 + 排序
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
    {"name": "流光", "mult": 30.0},
    {"name": "水晶", "mult": 20.0},
    {"name": "金", "mult": 10.0},
    {"name": "银", "mult": 3.0},
    {"name": "无", "mult": 1.0},
]
mutations_moon = [{"name": "星空", "mult": 40.0}] + mutations_earth

# 洒水器类型及对应的 k 值
sprinklers = [
    {"name": "空刷", "k": 1.0},
    {"name": "简易", "k": 1.2},
    {"name": "标准", "k": 1.7},
    {"name": "白银", "k": 2.4},
    {"name": "黄金", "k": 3.4},
]


def format_price(value):
    """格式化价格：≥10000 显示为 X.XX万，否则完整数字"""
    if value >= 10000:
        v = value / 10000
        return f"{v:.2f}万"
    else:
        return f"{int(round(value)):,}"


# 稀有度 → 背景色 映射（名称列专用）
rarity_colors = {
    "common": "#FFFFFF",  # 不染色
    "uncommon": "#78F983",  # 极浅绿
    "rare": "#75C3FB",  # 极浅蓝（天蓝系）
    "legendary": "#EC6FFF",  # 极浅紫
    "devine": "#FB9D81",  # 极浅橙
    "prismatic": "#FC6E43",  # 明显橙色（棱彩用较深一点区分）
}


def get_rarity(name: str) -> str:
    if name in common:
        return "common"
    if name in uncommon:
        return "uncommon"
    if name in rare:
        return "rare"
    if name in legendary:
        return "legendary"
    if name in devine:
        return "devine"
    if name in prismatic:
        return "prismatic"
    return "common"  # 默认


def create_sheet(
    workbook: xlsxwriter.Workbook,
    sheet_name: str,
    plants: list[dict],
    mutations: list[dict],
):
    worksheet = workbook.add_worksheet(sheet_name)

    headers = ["作物名称", "突变", "规模", "空刷", "简易", "标准", "白银", "黄金"]

    # 洒水器列背景色
    colors = ["#FFFFFF", "#D3D3D3", "#90EE90", "#87CEFA", "#FFCC66"]

    # 通用单元格格式（无背景）
    cell_fmt = workbook.add_format({"border": 1})
    cell_fmt_top = workbook.add_format({"border": 1, "top": 2})

    # 表头格式
    header_fmt = workbook.add_format(
        {"bold": True, "align": "center", "valign": "vcenter", "border": 1}
    )

    # 洒水器列格式（带背景）
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

    # 列宽设置
    worksheet.set_column(0, 0, 20)  # 作物名称
    worksheet.set_column(1, 1, 10)  # 突变
    worksheet.set_column(2, 2, 10)  # 规模
    worksheet.set_column(3, 7, 28)  # 洒水器列

    row = 1
    for plant in plants:
        name = plant["name"]
        rarity = get_rarity(name)
        bg_color = rarity_colors.get(rarity, "#FFFFFF")

        # 为名称列创建带颜色的格式（第一行粗上框，其余细框）
        name_fmt_base = {
            "bg_color": bg_color,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "border": 1,
        }
        name_fmt_top = workbook.add_format({**name_fmt_base, "top": 2})
        name_fmt_normal = workbook.add_format(name_fmt_base)

        max_weight = plant["maxWeight"]
        price_coeff = plant["priceCoefficient"]
        first_row_of_plant = True

        for scale, is_giant in [
            ("巨大", True),
            ("普通", False),
        ]:
            for mut in mutations:
                mult = mut["mult"]
                mut_name = mut["name"]

                # 名称列使用带颜色的格式
                name_fmt = name_fmt_top if first_row_of_plant else name_fmt_normal

                worksheet.write(row, 0, name, name_fmt)
                worksheet.write(
                    row, 1, mut_name, cell_fmt_top if first_row_of_plant else cell_fmt
                )
                worksheet.write(
                    row, 2, scale, cell_fmt_top if first_row_of_plant else cell_fmt
                )

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

    # 冻结前三列 + 第一行
    worksheet.freeze_panes(1, 3)


# 生成 Excel
workbook = xlsxwriter.Workbook("作物期望价值表.xlsx")
create_sheet(workbook, "地球", earth_plants, mutations_earth)
create_sheet(workbook, "月球", moon_plants, mutations_moon)
workbook.close()

print("✅ 已生成：作物期望价值表.xlsx")
print(" • 作物名称列按稀有度上色（同一作物所有行同色）")
print(" • 普通行 G=1，巨大行 G=5")
print(" • 每格为期望价值（已乘突变倍数）")
print(" • 颜色对应：普通白 / 不常见浅绿 / 稀有天蓝 / 传说浅紫 / 神圣浅橙 / 棱彩橙色")

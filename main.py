import json
import math
from typing import Iterable
import xlsxwriter

# 加载 JSON 数据
with open("plants.json", "r", encoding="utf-8") as f:
    plants_data = json.load(f)

# fmt: off
common = ["月光草", "灰壤豆", "土豆", "香菇", "番茄", "波斯菊"]
uncommon = ["月灯草", "月番茄", "大豆", "竹子", "黄瓜"]
rare = ["西瓜", "梨", "橘子", "玉米", "白菜", "牵牛花", "棉花", "月环树", "银灰苔"]
legendary = ["月莓", "星叶菜", "苹果", "石榴", "香蕉", "车厘子", "椰子", "南瓜"]
devine = ["草莓", "猕猴桃", "荔枝", "榴莲", "月核树", "液光藤"]
prismatic = [
    "月影梅","幻月花","星空玫瑰","红包树","月兔",
    "向日葵","松果","大王菊","葡萄","蟠桃",
    "惊奇菇","仙人掌象","魔鬼朝天椒"
]
# fmt: on


def sort_plants(plants: Iterable[dict]) -> list:
    return sorted(
        plants,
        key=lambda p: p["maxWeight"] ** 1.5 * p["priceCoefficient"],
        reverse=True,
    )


earth_plants = sort_plants(p for p in plants_data if p["type"] == "普通")
moon_plants = sort_plants(p for p in plants_data if p["type"] == "月球")


mutations_earth = [
    {"name": "流光", "mult": 30.0},
    {"name": "水晶", "mult": 20.0},
    {"name": "金", "mult": 10.0},
    {"name": "银", "mult": 3.0},
    {"name": "无", "mult": 1.0},
]

mutations_moon = [{"name": "星空", "mult": 40.0}] + mutations_earth


sprinklers = [
    {"name": "空刷", "k": 1.0},
    {"name": "简易", "k": 1.2},
    {"name": "标准", "k": 1.7},
    {"name": "白银", "k": 2.4},
    {"name": "黄金", "k": 3.4},
]


def format_price(v: float):
    if v >= 10000:
        return f"{v / 10000:.2f}万"
    return f"{int(round(v)):,}"


rarity_colors = {
    "common": "#FFFFFF",
    "uncommon": "#78F983",
    "rare": "#75C3FB",
    "legendary": "#EC6FFF",
    "devine": "#FB9D81",
    "prismatic": "#FC6E43",
}


def get_rarity(name: str):
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
    return "common"


def create_formats(
    workbook: xlsxwriter.Workbook,
    colors: list[str],
):

    formats = {}

    formats["cell"] = workbook.add_format(
        {"border": 1, "align": "center", "valign": "vcenter"}
    )
    formats["cell_top"] = workbook.add_format(
        {"border": 1, "top": 2, "align": "center", "valign": "vcenter"}
    )
    formats["cell_dash"] = workbook.add_format(
        {"border": 1, "top": 8, "align": "center", "valign": "vcenter"}
    )

    formats["sprinkler"] = []
    formats["sprinkler_top"] = []
    formats["sprinkler_dash"] = []

    for c in colors:
        formats["sprinkler"].append(
            workbook.add_format(
                {"border": 1, "bg_color": c, "align": "center", "valign": "vcenter"}
            )
        )

        formats["sprinkler_top"].append(
            workbook.add_format(
                {
                    "border": 1,
                    "top": 2,
                    "bg_color": c,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
        )

        formats["sprinkler_dash"].append(
            workbook.add_format(
                {
                    "border": 1,
                    "top": 8,
                    "bg_color": c,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
        )

    formats["header"] = workbook.add_format(
        {"bold": True, "border": 1, "align": "center", "valign": "vcenter"}
    )

    return formats


def create_sheet(
    workbook: xlsxwriter.Workbook,
    sheet_name: str,
    plants: list[dict],
    mutations: list[dict],
    expection: bool = True,
):

    worksheet = workbook.add_worksheet(sheet_name)

    headers = ["作物名称", "突变", "规模", "空刷", "简易", "标准", "白银", "黄金"]

    colors = ["#FFFFFF", "#D3D3D3", "#90EE90", "#87CEFA", "#FFCC66"]

    formats = create_formats(workbook, colors)

    for col, h in enumerate(headers):
        worksheet.write(0, col, h, formats["header"])

    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 2, 10)
    worksheet.set_column(3, 7, 18)

    row = 1

    const = (2 / 5) * (4 * math.sqrt(2) - 1)

    for plant in plants:
        name = plant["name"]
        rarity = get_rarity(name)

        name_fmt = workbook.add_format(
            {
                "border": 1,
                "bg_color": rarity_colors[rarity],
                "align": "center",
                "valign": "vcenter",
            }
        )

        name_fmt_top = workbook.add_format(
            {
                "border": 1,
                "top": 2,
                "bg_color": rarity_colors[rarity],
                "align": "center",
                "valign": "vcenter",
            }
        )

        name_fmt_dash = workbook.add_format(
            {
                "border": 1,
                "top": 8,
                "bg_color": rarity_colors[rarity],
                "align": "center",
                "valign": "vcenter",
            }
        )

        max_weight = plant["maxWeight"]
        price_coeff = plant["priceCoefficient"]

        first_row = True

        for scale, is_giant in [("巨大", True), ("普通", False)]:
            for i, mut in enumerate(mutations):
                mut_name = mut["name"]
                mult = mut["mult"]

                is_dash_line = scale == "普通" and i == 0

                if first_row:
                    fmt = formats["cell_top"]
                    name_format = name_fmt_top
                    sprinkler_fmt = formats["sprinkler_top"]
                elif is_dash_line:
                    fmt = formats["cell_dash"]
                    name_format = name_fmt_dash
                    sprinkler_fmt = formats["sprinkler_dash"]
                else:
                    fmt = formats["cell"]
                    name_format = name_fmt
                    sprinkler_fmt = formats["sprinkler"]

                worksheet.write(row, 0, name, name_format)
                worksheet.write(row, 1, mut_name, fmt)
                worksheet.write(row, 2, scale, fmt)

                G = 5.0 if is_giant else 1.0

                for i, sp in enumerate(sprinklers):
                    k = sp["k"]
                    if expection:
                        effective_w = max_weight * k * G / 34
                        expected = const * price_coeff * (effective_w**1.5) * mult
                        worksheet.write(
                            row, i + 3, format_price(expected), sprinkler_fmt[i]
                        )
                    else:
                        min_weight = max_weight * k * G / 34
                        max_weight = min_weight * 2
                        price_min = price_coeff * (min_weight**1.5) * mult
                        price_max = price_coeff * (max_weight**1.5) * mult
                        worksheet.write(
                            row,
                            i + 3,
                            f"{format_price(price_min)}~{format_price(price_max)}",
                            sprinkler_fmt[i],
                        )

                first_row = False
                row += 1

    worksheet.freeze_panes(1, 3)


OUT_FILE = "作物洒水价值表.xlsx"
workbook = xlsxwriter.Workbook(OUT_FILE)

create_sheet(workbook, "地球期望", earth_plants, mutations_earth)
create_sheet(workbook, "地球范围", earth_plants, mutations_earth, expection=False)
create_sheet(workbook, "月球期望", moon_plants, mutations_moon)
create_sheet(workbook, "月球范围", moon_plants, mutations_moon, expection=False)

workbook.close()

print(f"✅ 已生成：{OUT_FILE}")

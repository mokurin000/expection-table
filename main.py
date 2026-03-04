import json
import xlsxwriter
import math

# 加载 JSON 数据
with open('plants.json', 'r', encoding='utf-8') as f:
    plants_data = json.load(f)

# 分离地球和月球作物
earth_plants = [p for p in plants_data if p['type'] == '普通']
moon_plants = [p for p in plants_data if p['type'] == '月球']

# 基础突变类型及倍数
mutations_earth = [
    {'name': '无', 'mult': 1.0},
    {'name': '银', 'mult': 3.0},
    {'name': '金', 'mult': 10.0},
    {'name': '水晶', 'mult': 20.0},
    {'name': '流光', 'mult': 30.0},
]

mutations_moon = mutations_earth + [
    {'name': '星空', 'mult': 40.0}
]

# 洒水器类型及对应的 k 值（空刷=1, 简易=1.2, 标准=1.7, 白银=2.4, 黄金=3.4）
sprinklers = [
    {'name': '空刷', 'k': 1.0},
    {'name': '简易', 'k': 1.2},
    {'name': '标准', 'k': 1.7},
    {'name': '白银', 'k': 2.4},
    {'name': '黄金', 'k': 3.4},
]

def format_price(value):
    """格式化价格：≥10000 显示为 X.X万（保留1位小数），否则完整数字"""
    if value >= 10000:
        v = round(value / 10000, 1)
        return f"{v:.1f}万"
    else:
        return f"{int(round(value)):,}"

def create_sheet(workbook, sheet_name, plants, mutations):
    worksheet = workbook.add_worksheet(sheet_name)
    
    # 表头（格式不变）
    headers = ['作物名称', '突变', '规模', '空刷', '简易', '标准', '白银', '黄金']
    
    # 列背景色：白 灰 绿 蓝 橙（仅洒水器5列）
    colors = ['#FFFFFF', '#D3D3D3', '#90EE90', '#87CEFA', '#FFCC66']
    col_formats = [workbook.add_format({'bg_color': c, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True}) for c in colors]
    
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    
    # 写入表头
    for col, header in enumerate(headers):
        if col < 3:
            worksheet.write(0, col, header, header_fmt)
        else:
            h_fmt = workbook.add_format({
                'bg_color': colors[col-3],
                'bold': True,
                'align': 'center',
                'valign': 'vcenter'
            })
            worksheet.write(0, col, header, h_fmt)
    
    # 列宽设置
    worksheet.set_column(0, 0, 20)   # 作物名称
    worksheet.set_column(1, 1, 10)   # 突变
    worksheet.set_column(2, 2, 10)   # 规模（普通/巨大）
    worksheet.set_column(3, 7, 28)   # 洒水器列（期望值）
    
    row = 1
    for plant in plants:
        name = plant['name']
        max_weight = plant['maxWeight']
        price_coeff = plant['priceCoefficient']
        
        for mut in mutations:
            mult = mut['mult']
            mut_name = mut['name']
            
            for scale, is_giant in [('普通', False), ('巨大', True)]:
                worksheet.write(row, 0, name)
                worksheet.write(row, 1, mut_name)
                worksheet.write(row, 2, scale)
                
                G = 5.0 if is_giant else 1.0
                
                for s_idx, sp in enumerate(sprinklers):
                    k = sp['k']
                    
                    # 期望价值公式
                    effective_w = max_weight * k * G / 34
                    const = (2 / 5) * (4 * math.sqrt(2) - 1)
                    expected = const * price_coeff * (effective_w ** 1.5) * mult
                    
                    cell_text = format_price(expected)
                    worksheet.write(row, s_idx + 3, cell_text, col_formats[s_idx])
                
                row += 1
    
    # 冻结前三列（作物名称、突变、规模）
    worksheet.freeze_panes(1, 3)

# 生成 Excel
workbook = xlsxwriter.Workbook('作物期望价值表_分行版.xlsx')
create_sheet(workbook, '地球', earth_plants, mutations_earth)
create_sheet(workbook, '月球', moon_plants, mutations_moon)
workbook.close()

print("✅ 已生成：作物期望价值表_分行版.xlsx")
print("   • 格式与上次完全一致（作物名称 | 突变 | 规模 | 五列洒水器）")
print("   • 每格现在显示的是使用公式计算的**期望价值**（已乘突变倍数）")
print("   • 普通行 G=1，巨大行 G=5")
print("   • 大额自动转为 X.X万（保留1位小数）")
print("   • 可直接筛选/排序/筛选巨大化或特定突变")

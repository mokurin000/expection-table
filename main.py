import json
import xlsxwriter

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

# 洒水器配置（百分比已 ×100）
sprinklers = [
    {'name': '空刷',   'min': 2.94,  'max': 5.88,   'minG': 14.71, 'maxG': 29.41},
    {'name': '简易',   'min': 3.53,  'max': 7.06,   'minG': 17.65, 'maxG': 35.29},
    {'name': '标准',   'min': 5.00,  'max': 10.00,  'minG': 25.00, 'maxG': 50.00},
    {'name': '白银',   'min': 7.06,  'max': 14.12,  'minG': 35.29, 'maxG': 70.59},
    {'name': '黄金',   'min': 10.00, 'max': 20.00,  'minG': 50.00, 'maxG': 100.00},
]

def format_price(value):
    """格式化价格：≥10000 显示为 X.X万（保留1位小数），否则完整数字"""
    if value >= 10000:
        v = round(value / 10000, 1)          # 四舍五入到1位小数
        return f"{v:.1f}万"                  # 强制显示1位小数（如 1.0万）
    else:
        return f"{int(round(value)):,}"

def create_sheet(workbook, sheet_name, plants, mutations):
    worksheet = workbook.add_worksheet(sheet_name)
    
    # 表头
    headers = ['作物名称', '空刷', '简易', '标准', '白银', '黄金']
    
    # 列背景色：白 灰 绿 蓝 橙
    colors = ['#FFFFFF', '#D3D3D3', '#90EE90', '#87CEFA', '#FFCC66']
    col_formats = [workbook.add_format({'bg_color': c, 'align': 'center', 'valign': 'vcenter'}) for c in colors]
    
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    
    # 写入表头
    for col, header in enumerate(headers):
        if col == 0:
            worksheet.write(0, col, header, header_fmt)
        else:
            h_fmt = workbook.add_format({
                'bg_color': colors[col-1],
                'bold': True,
                'align': 'center',
                'valign': 'vcenter'
            })
            worksheet.write(0, col, header, h_fmt)
    
    # 列宽（加大一点容纳“X.X万 ~ Y.Y万”）
    worksheet.set_column(0, 0, 28)
    worksheet.set_column(1, 5, 38)
    
    row = 1
    for plant in plants:
        name = plant['name']
        max_weight = plant['maxWeight']
        price_coeff = plant['priceCoefficient']
        
        for mut in mutations:
            display_name = f"{name} ({mut['name']})"
            worksheet.write(row, 0, display_name)
            
            for s_idx, sp in enumerate(sprinklers):
                mult = mut['mult']
                
                # 非巨大化范围
                w_min = max_weight * (sp['min'] / 100)
                w_max = max_weight * (sp['max'] / 100)
                p_min_norm = (w_min ** 1.5) * price_coeff * mult
                p_max_norm = (w_max ** 1.5) * price_coeff * mult
                
                # 巨大化范围
                w_min_g = max_weight * (sp['minG'] / 100)
                w_max_g = max_weight * (sp['maxG'] / 100)
                p_min_g = (w_min_g ** 1.5) * price_coeff * mult
                p_max_g = (w_max_g ** 1.5) * price_coeff * mult
                
                # 格式化
                norm_range = f"{format_price(p_min_norm)} ~ {format_price(p_max_norm)}"
                g_range   = f"{format_price(p_min_g)} ~ {format_price(p_max_g)}"
                
                cell_text = f"普通: {norm_range}\r\n巨大: {g_range}"
                
                worksheet.write(row, s_idx + 1, cell_text, col_formats[s_idx])
            
            row += 1
    
    # 冻结首行 + 首列（作物名）
    worksheet.freeze_panes(1, 1)

# 生成 Excel
workbook = xlsxwriter.Workbook('作物底价范围表_小数万.xlsx')
create_sheet(workbook, '地球', earth_plants, mutations_earth)
create_sheet(workbook, '月球', moon_plants, mutations_moon)
workbook.close()

print("已生成：作物底价范围表_小数万.xlsx")
print("大额价值示例：")
print("  12345 → 1.2万")
print("  10000 → 1.0万")
print("  56789 → 5.7万")
print("  9999  → 9,999（不变）")

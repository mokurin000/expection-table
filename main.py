import json
import xlsxwriter

with open('plants.json', 'r', encoding='utf-8') as f:
    plants_data = json.load(f)

earth_plants = [p for p in plants_data if p['type'] == '普通']
moon_plants = [p for p in plants_data if p['type'] == '月球']

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

sprinklers = [
    {'name': '空刷',   'min': 2.94,  'max': 5.88,   'minG': 14.71, 'maxG': 29.41},
    {'name': '简易',   'min': 3.53,  'max': 7.06,   'minG': 17.65, 'maxG': 35.29},
    {'name': '标准',   'min': 5.00,  'max': 10.00,  'minG': 25.00, 'maxG': 50.00},
    {'name': '白银',   'min': 7.06,  'max': 14.12,  'minG': 35.29, 'maxG': 70.59},
    {'name': '黄金',   'min': 10.00, 'max': 20.00,  'minG': 50.00, 'maxG': 100.00},
]

def format_price(value):
    if value >= 10000:
        v = round(value / 10000, 1)
        return f"{v:.1f}万"
    else:
        return f"{int(round(value)):,}"

def create_sheet(workbook, sheet_name, plants, mutations):
    worksheet = workbook.add_worksheet(sheet_name)
    
    headers = ['作物名称', '突变', '规模', '空刷', '简易', '标准', '白银', '黄金']
    
    colors = ['#FFFFFF', '#D3D3D3', '#90EE90', '#87CEFA', '#FFCC66']
    col_formats = [workbook.add_format({'bg_color': c, 'align': 'center', 'valign': 'vcenter'}) for c in colors]
    
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    
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
    
    worksheet.set_column(0, 0, 20)     # 作物名称
    worksheet.set_column(1, 1, 10)     # 突变
    worksheet.set_column(2, 2, 10)     # 规模（普通/巨大）
    worksheet.set_column(3, 7, 32)     # 洒水器列
    
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
                
                for s_idx, sp in enumerate(sprinklers):
                    if is_giant:
                        pct_min = sp['minG']
                        pct_max = sp['maxG']
                    else:
                        pct_min = sp['min']
                        pct_max = sp['max']
                    
                    w_min = max_weight * (pct_min / 100)
                    w_max = max_weight * (pct_max / 100)
                    p_min = (w_min ** 1.5) * price_coeff * mult
                    p_max = (w_max ** 1.5) * price_coeff * mult
                    
                    range_text = f"{format_price(p_min)} ~ {format_price(p_max)}"
                    worksheet.write(row, s_idx + 3, range_text, col_formats[s_idx])
                
                row += 1
    
    worksheet.freeze_panes(1, 3)  # 冻结前三列（作物、突变、规模）

workbook = xlsxwriter.Workbook('作物底价范围表_分行版.xlsx')
create_sheet(workbook, '地球', earth_plants, mutations_earth)
create_sheet(workbook, '月球', moon_plants, mutations_moon)
workbook.close()

print("已生成：作物底价范围表_分行版.xlsx")
print("现在每种作物+突变+规模（普通/巨大）独立一行，便于筛选、排序、过滤。")
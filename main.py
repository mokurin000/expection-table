import json
import xlsxwriter

# 加载 JSON 数据
with open('plants.json', 'r', encoding='utf-8') as f:
    plants_data = json.load(f)

# 分离地球和月球作物
earth_plants = [p for p in plants_data if p['type'] == '普通']
moon_plants = [p for p in plants_data if p['type'] == '月球']

# 突变类型及倍数（基础突变）
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

# 洒水器类型及重量百分比（已从图片表中提取）
sprinklers = [
    {'name': '空刷', 'min_pct': 2.94, 'max_pct': 5.88, 'minG_pct': 14.71, 'maxG_pct': 29.41},
    {'name': '简易', 'min_pct': 3.53, 'max_pct': 7.06, 'minG_pct': 17.65, 'maxG_pct': 35.29},
    {'name': '标准', 'min_pct': 5.0, 'max_pct': 10.0, 'minG_pct': 25.0, 'maxG_pct': 50.0},
    {'name': '白银', 'min_pct': 7.06, 'max_pct': 14.12, 'minG_pct': 35.29, 'maxG_pct': 70.59},
    {'name': '黄金', 'min_pct': 10.0, 'max_pct': 20.0, 'minG_pct': 50.0, 'maxG_pct': 100.0},
]

def create_sheet(workbook, sheet_name, plants, mutations):
    worksheet = workbook.add_worksheet(sheet_name)
    
    # 表头
    headers = ['作物名称', '空刷', '简易', '标准', '白银', '黄金']
    
    # 右五列背景色（白、灰、绿、蓝、橙）
    colors = ['#FFFFFF', '#D3D3D3', '#90EE90', '#87CEFA', '#FFCC66']
    col_formats = [
        workbook.add_format({'bg_color': colors[i], 'align': 'center', 'valign': 'vcenter'})
        for i in range(5)
    ]
    
    # 表头格式（加粗）
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
    
    # 写入表头
    for col, header in enumerate(headers):
        if col == 0:
            worksheet.write(0, col, header, header_fmt)
        else:
            h_fmt = workbook.add_format({
                'bg_color': colors[col - 1],
                'bold': True,
                'align': 'center',
                'valign': 'vcenter'
            })
            worksheet.write(0, col, header, h_fmt)
    
    # 设置列宽
    worksheet.set_column(0, 0, 28)   # 作物名称列
    worksheet.set_column(1, 5, 32)   # 洒水器范围列
    
    # 写入数据
    row = 1
    for plant in plants:
        name = plant['name']
        max_weight = plant['maxWeight']
        price_coeff = plant['priceCoefficient']
        
        for mut in mutations:
            # 作物名称 + 突变类型（便于区分多行）
            display_crop = f"{name} ({mut['name']})"
            worksheet.write(row, 0, display_crop)
            
            for s_idx, sp in enumerate(sprinklers):
                # 计算该洒水器下非巨大化/巨大化的整体重量百分比范围
                min_p = min(sp['min_pct'], sp['minG_pct'])
                max_p = max(sp['max_pct'], sp['maxG_pct'])
                
                low_weight = max_weight * (min_p / 100)
                high_weight = max_weight * (max_p / 100)
                
                # 价值公式：(重量)^1.5 * 价格系数 * 突变倍数
                low_price = (low_weight ** 1.5) * price_coeff * mut['mult']
                high_price = (high_weight ** 1.5) * price_coeff * mut['mult']
                
                # 格式化范围（取整，带千位分隔符）
                range_text = f"{int(round(low_price)):,} - {int(round(high_price)):,}"
                
                worksheet.write(row, s_idx + 1, range_text, col_formats[s_idx])
            
            row += 1
    
    # 冻结首行
    worksheet.freeze_panes(1, 0)

# 生成 Excel
workbook = xlsxwriter.Workbook('作物底价范围表.xlsx')
create_sheet(workbook, '地球', earth_plants, mutations_earth)
create_sheet(workbook, '月球', moon_plants, mutations_moon)
workbook.close()

print("✅ Excel 文件已成功生成：作物底价范围表.xlsx")
print("   - 地球作物 Sheet：每种作物按「无、银、金、水晶、流光」5 行")
print("   - 月球作物 Sheet：每种作物按「无、银、金、水晶、流光、星空」6 行")
print("   - 每格显示该洒水器 + 该突变下的底价范围（min - max）")
print("   - 右五列已按要求设置背景色：白、灰、绿、蓝、橙")
from openpyxl import Workbook 
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side 
from openpyxl.utils import get_column_letter 

# 创建工作簿
wb = Workbook()
ws = wb.active
ws.title = "现金流量表"

# 设置表头
ws['A1'] = "[公司全称]"
ws['A2'] = "现金流量表"
ws['A3'] = "截至2025年12月31日止年度"

# 合并表头
ws.merge_cells('A1:C1')
ws.merge_cells('A2:C2')
ws.merge_cells('A3:C3')

# 设置表头样式
header_font = Font(name='微软雅黑', size=14, bold=True)
for cell in ws['A1:C3']:
    for c in cell:
        c.font = header_font
        c.alignment = Alignment(horizontal='center', vertical='center')

# 设置列宽
ws.column_dimensions['A'].width = 45
ws.column_dimensions['B'].width = 15
ws.column_dimensions['C'].width = 15

# 定义项目列表
items = [
    ("一、经营活动产生的现金流量", None, None),
    ("销售商品、提供劳务收到的现金", "C5", "D5"),
    ("收到的税费返还", "C6", "D6"),
    ("收到其他与经营活动有关的现金", "C7", "D7"),
    ("经营活动现金流入小计", "C8", "D8"),
    ("购买商品、接受劳务支付的现金", "C11", "D11"),
    ("支付给职工以及为职工支付的现金", "C12", "D12"),
    ("支付的各项税费", "C13", "D13"),
    ("支付其他与经营活动有关的现金", "C14", "D14"),
    ("经营活动现金流出小计", "C15", "D15"),
    ("经营活动产生的现金流量净额", "C16", "D16"),
    ("二、投资活动产生的现金流量", None, None),
    ("收回投资收到的现金", "C19", "D19"),
    ("取得投资收益收到的现金", "C20", "D20"),
    ("处置固定资产、无形资产和其他长期资产收回的现金净额", "C21", "D21"),
    ("处置子公司及其他营业单位收到的现金净额", "C22", "D22"),
    ("收到其他与投资活动有关的现金", "C23", "D23"),
    ("投资活动现金流入小计", "C24", "D24"),
    ("购建固定资产、无形资产和其他长期资产支付的现金", "C27", "D27"),
    ("投资支付的现金", "C28", "D28"),
    ("取得子公司及其他营业单位支付的现金净额", "C29", "D29"),
    ("支付其他与投资活动有关的现金", "C30", "D30"),
    ("投资活动现金流出小计", "C31", "D31"),
    ("投资活动产生的现金流量净额", "C32", "D32"),
    ("三、筹资活动产生的现金流量", None, None),
    ("吸收投资收到的现金", "C35", "D35"),
    ("取得借款收到的现金", "C36", "D36"),
    ("发行债券收到的现金", "C37", "D37"),
    ("收到其他与筹资活动有关的现金", "C38", "D38"),
    ("筹资活动现金流入小计", "C39", "D39"),
    ("偿还债务支付的现金", "C42", "D42"),
    ("分配股利、利润或偿付利息支付的现金", "C43", "D43"),
    ("支付其他与筹资活动有关的现金", "C44", "D44"),
    ("筹资活动现金流出小计", "C45", "D45"),
    ("筹资活动产生的现金流量净额", "C46", "D46"),
    ("四、汇率变动对现金及现金等价物的影响", "C47", "D47"),
    ("五、现金及现金等价物净增加额", "C48", "D48"),
    ("加：期初现金及现金等价物余额", "C49", "D49"),
    ("六、期末现金及现金等价物余额", "C50", "D50")
]

# --- 填充项目逻辑优化 ---
# --- 3. 填充逻辑 (核心修正区) ---
row = 5
indent_style = Alignment(horizontal='left', vertical='center', indent=2)

for item, col_b, col_c in items:
    if item is None:
        row += 1
        continue
    
    ws[f'A{row}'] = item
    # 设置缩进
    if not any(item.startswith(x) for x in ["一、", "二、", "三、", "四、", "五、", "六、"]):
        ws[f'A{row}'].alignment = indent_style

    if col_b:
        # 提取字母用于拼公式 (例如 "C5" -> "C")
        col_b_let = "".join(filter(str.isalpha, col_b))
        col_c_let = "".join(filter(str.isalpha, col_c))
        
        # 初始化为0，防止 #VALUE!
        ws[col_b] = 0
        ws[col_c] = 0
        
        if "小计" in item:
            start_r = row - 3 if "流入" in item else row - 4
            ws[col_b] = f"=SUM({col_b_let}{start_r}:{col_b_let}{row-1})"
            ws[col_c] = f"=SUM({col_c_let}{start_r}:{col_c_let}{row-1})"
            
        elif "净额" in item:
            if "经营" in item: r1, r2 = 8, 15
            elif "投资" in item: r1, r2 = 24, 31
            else: r1, r2 = 39, 45 # 筹资
            
            if "净增加额" not in item:
                # 注意：这里必须用 ws[col_b] 而不是 ws[col_b_let]
                ws[col_b] = f"={col_b_let}{r1}-{col_b_let}{r2}"
                ws[col_c] = f"={col_c_let}{r1}-{col_c_let}{r2}"
            else:
                ws[col_b] = "=C16+C32+C46+C47"
                ws[col_c] = "=D16+D32+D46+D47"
        
        elif "期末" in item:
            ws[col_b] = "=C48+C49"
            ws[col_c] = "=D48+D49"
            
    row += 1

# --- 样式与格式优化 ---
gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

for r in range(5, 51):
    cell_a = ws[f'A{r}']
    if not cell_a.value: continue
    
    # 统一数字格式为 C 列和 D 列
    for col_let in ['C', 'D']:
        cell = ws[f'{col_let}{r}']
        cell.number_format = '#,##0.00'
        
        # 自动变色逻辑
        if any(keyword in cell_a.value for keyword in ["流入", "收到"]):
            cell.font = Font(color="006400") # 深绿
        elif any(keyword in cell_a.value for keyword in ["流出", "支付", "偿还"]):
            cell.font = Font(color="FF0000") # 红色
            
        # 汇总行底色
        if any(keyword in cell_a.value for keyword in ["小计", "净额", "增加额", "余额"]):
            cell.fill = gray_fill
            cell.font = Font(bold=True)

# 边框与保存
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
for row_cells in ws['A5:D50']:
    for cell in row_cells:
        cell.border = thin_border

wb.save("现金流量表_中国会计准则_2025版.xlsx")
print("✅ 报表已更新，请查看文件。")

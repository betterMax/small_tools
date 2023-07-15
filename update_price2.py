import xlwings as xw
from get_latest_price import get_latest_price

# 加载工作簿
wb = xw.Book('Resource/test.xlsx')

# 选择工作表
ws = wb.sheets['期货持仓']  # 根据实际的sheet名称修改

# 读取单元格
code = ws.range('A2').value
print(f'update {code} price')
latest_price = get_latest_price(code)
# 修改单元格
# ws['Q621'] = latest_price
# 尝试将价格转换为浮点数
try:
    price = float(latest_price)
except ValueError:
    print(f'Warning: Cannot convert price "{latest_price}" to number. Please check the price.')
else:
    # 如果转换成功，更新价格
    ws.range('M2').value = price

# 保存修改
wb.save('Resource/test.xlsx')

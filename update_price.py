from openpyxl import load_workbook
from get_latest_price import get_latest_price
import time

# 加载工作簿
wb = load_workbook(filename='Resource/target.xlsx', data_only=True)

mode = 'work'
# 待处理的sheet列表
if mode == 'work':
    sheets = ['判断', '实际']
else:
    sheets = ['判断', '实际', '虚拟']

for sheet in sheets:
    # 选择工作表
    ws = wb[sheet]

    # 初始化行号
    row = 2
    
    # 循环处理每一行，直到A列没有数据
    while ws[f'A{row}'].value is not None:
        # 读取单元格
        code = ws[f'A{row}'].value
        print(f'update {code} price')
        latest_price = get_latest_price(code)

        # 尝试将价格转换为浮点数
        try:
            price = float(latest_price)
        except ValueError:
            print(f'Warning: Cannot convert price "{latest_price}" to number. Please check the price.')
        else:
            # 如果转换成功，更新价格
            ws[f'B{row}'].value = price

        # 处理下一行
        row += 1

        # 暂停一段时间
        time.sleep(5)  # 这里设置暂停一秒，可以根据实际需要调整


# 保存修改
wb.save('Resource/target.xlsx')

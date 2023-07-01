import pandas as pd


def excel_column_to_list(file_path, column_name):
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 获取指定列的数据
    column_data = df[column_name].tolist()

    # 在每个元素外添加单引号
    # column_data = ["'" + str(item) + "'" for item in column_data]

    return column_data

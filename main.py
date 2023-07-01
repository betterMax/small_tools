from utilities import excel_column_to_list


def main():
    # 获取用户输入的功能选择
    choice = input("请选择要执行的功能（1-导出Excel列数据为列表）：")

    if choice == '1':
        # 获取用户输入的文件路径和列名
        file_path = input("请输入文件路径：")
        column_name = input("请输入列名：")

        # 调用函数并获取结果
        result = excel_column_to_list(file_path, column_name)

        # 打印结果
        print(result)
    else:
        print("无效的选择")


if __name__ == '__main__':
    main()


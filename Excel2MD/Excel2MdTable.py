# -*- coding: utf-8 -*-
# @Time    : 2023/4/12 16:27
# @Author  : yan.wei
# @Email   : 13675196684@163.com
# @File    : Excel2MdTable.py
# @Software: PyCharm


# 将Excel文件转换为Markdown表格的程序。它使用了pandas库来读取Excel文件，并将其转换为DataFrame对象。然后，它使用DataFrame对象的列名和值来构建Markdown表格。具体来说，它首先使用列名构建表格的标题和分隔线，然后使用每行的值构建表格的行。最后，它将结果打印到控制台。

# 将excel文件转换为markdown表格
import pandas as pd

class ToMd():
    def __init__(self):
        pass

    def excel2Md(self, path, sheetname):
        # 读取Excel文件
        df = pd.read_excel(path, sheet_name=sheetname)
        # 构建表格标题
        title = " | "
        # 构建表格分隔线
        spiltLine = " | "
        # 遍历DataFrame对象的列名
        for i in df.columns.values:
            # 将列名添加到标题中
            title += i + " | "
            # 将分隔线添加到分隔线中
            spiltLine += " --- | "
        # 打印表格标题
        print(title)
        # 打印表格分隔线
        print(spiltLine)
        # 遍历DataFrame对象的每一行
        for i in range(len(df)):
            # 构建表格行
            line = " | "
            # 遍历DataFrame对象的列名
            for j in df.columns.values:
                # 将单元格的值添加到行中
                line += str(df.loc[i, j]) + " | "
            # 打印表格行，并将"nan"替换为"-"
            print(line.replace("nan", " - "))

if __name__ == "__main__":
    # 创建ToMd对象
    tomd = ToMd()
    # 调用excel2Md方法，将Excel文件转换为Markdown表格
    tomd.excel2Md("test.xlsx", "Sheet1")



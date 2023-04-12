# -*- coding: utf-8 -*-
# @Time    : 2023/4/12 16:27
# @Author  : yan.wei
# @Email   : 13675196684@163.com
# @File    : Excel2MdTable.py
# @Software: PyCharm


# 将excel文件转换为markdown表格
import pandas as pd

class ToMd():
    def __init__(self, path, sheetname):
        self.path = path
        self.sheetname = sheetname

    def excel2Md(self, path, sheetname):
        df = pd.read_excel(path, sheet_name=sheetname)
        title = " | "
        spiltLine = " | "
        for i in df.columns.values:
            title += i + " | "
            spiltLine += " --- | "
        print(title)
        print(spiltLine)
        for i in range(len(df)):
            line = " | "
            for j in df.columns.values:
                line += str(df.loc[i, j]) + " | "
            print(line.replace("nan", " - "))

if __name__ == "__main__":
    tomd = ToMd()
    tomd.excel2Md("test.xlsx", "Sheet1")



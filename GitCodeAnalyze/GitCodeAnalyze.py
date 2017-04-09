#! /usr/bin/env python

# -*- coding: utf-8 -*-

import os
import xlrd
import xlwt


def open_excel(file='GitResult.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))


def excel_table_by_name(file='GitResult.xls', top_index=0, by_name=u'Sheet1'):
    author = []
    since = []
    until = []
    git_dir = []
    branch = []
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    n_rows = table.nrows  # 行数
    top_data = table.row_values(top_index)  # 某一行数据
    if not len(top_data) >= 5:
        print("ERROR, Excel row_values are wrong!")
        return

    for row_index in range(1, n_rows):
        row = table.row_values(row_index)
        if len(row) >= 5:
            author.append(row[0])
            since.append(row[1])
            until.append(row[2])
            git_dir.append(row[3])
            branch.append(row[4])
        else:
            author.append('')
            print("ERROR, Please Check Excel")

    return author, since, until, git_dir, branch


def git_code_analyze():
    author, since, until, git_dir, branch = excel_table_by_name('GitResult.xls', 0, 'Sheet1')
    if all(author) and all(since) and all(until) and all(git_dir) and all(branch):
        pass
    else:
        print("ERROR, Data not enough")
        return


if __name__ == "__main__":
    git_code_analyze()
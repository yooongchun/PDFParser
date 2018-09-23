# -*- coding:utf-8 -*-

__author__ = "yooongchun"
__email__ = "yooongchun@foxmail.com"
__wechat__ = '18217235290'

import re
import xlrd
from xlutils.copy import copy
import os
import xlwt


# load demo Excel
def load_demo_Excel(path, sheet_name):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name(sheet_name)
    rows = sheet.nrows
    info = []
    for row in range(rows):
        if row == 0:
            continue
        fund_name = sheet.cell_value(row, 0)
        report_date = sheet.cell_value(row, 1)
        report_date = xlrd.xldate_as_tuple(report_date, 0)
        report_date = str(report_date[0]) + '/' + str(
            report_date[1]) + '/' + str(report_date[2])
        filing_date = sheet.cell_value(row, 2)
        filing_date = xlrd.xldate_as_tuple(filing_date, 0)
        filing_date = str(filing_date[0]) + '/' + str(
            filing_date[1]) + '/' + str(filing_date[2])
        url = sheet.cell_value(row, 3)
        info.append({
            'fund_name': str(fund_name),
            'report_date': str(report_date),
            'filing_date': str(filing_date),
            'url': url
        })
    return info


def load_Excel(path):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name("Sheet1")
    rows = sheet.nrows
    info = []
    for row in range(rows):
        row_value = sheet.row_values(row)
        info.append(row_value)
    return info


def add_url(demo, info, info_path):
    for one in demo:
        date = one['report_date'].split('/')
        base_name = one['fund_name'] + '_' + date[0] + date[1] + '.xls'
        if os.path.basename(info_path) == base_name:
            for i in range(len(info)):
                info[i].append(one['url'])
    return info


def load_folder(folder):
    files = os.listdir(folder)
    paths = []
    for file in files:
        path = os.path.join(folder, file)
        if os.path.isfile(path) and os.path.splitext(path)[1] == '.xls':
            paths.append(path)
    return paths


def save(info, path, title_flag):
    # Header
    if title_flag:
        book = xlwt.Workbook()
        sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
        # Header
        sheet.write(0, 0, 'Fund_series')
        sheet.write(0, 1, 'Period of Report')
        sheet.write(0, 2, 'Filing Date')
        sheet.write(0, 3, 'Fund-Name')
        sheet.write(0, 4, 'Type1')
        sheet.write(0, 5, 'Type2')
        sheet.write(0, 6, 'Type3')
        sheet.write(0, 7, 'Stock')
        sheet.write(0, 8, 'Url')
        book.save(path)
    # Content
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name('Sheet1')
    rows = sheet.nrows
    copy_book = copy(book)
    sheet_copy = copy_book.get_sheet('Sheet1')
    for ind, one in enumerate(info):
        if ind == 0:
            continue
        for i in range(9):
            sheet_copy.write(rows + ind, i, one[i])
    copy_book.save(path)


if __name__ == '__main__':
    # paras
    base_dir = r'C:\Users\fanyu\Desktop\Project\PaidDevelopment\09_Spider_HTML_done_￥400\Excels'
    demo_path = r'C:\Users\fanyu\Desktop\Project\PaidDevelopment\09_Spider_HTML_done_￥400\List.xlsx'
    out_path = 'total.xls'

    # running...
    demo = load_demo_Excel(path=demo_path, sheet_name='webpage')
    files = load_folder(base_dir)
    for index, file in enumerate(files):
        print('progress:', index, '/', len(files))
        save(add_url(demo, load_Excel(file), file), out_path, True if index == 0 else False)

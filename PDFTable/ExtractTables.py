# -*- coding:utf-8 -*-

import pdfplumber
import re
import xlrd
import os
import xlwt
import threading
import multiprocessing
import shutil
from datetime import datetime

__author__ = 'yooongchun'
__email__ = 'yooongchun@foxmail.com'

'''
This program is designed for extracting specific table and text from PDFs.
@version: v1
@date:2018.09.26

'''


# 该类用来实现PDF表格和文字内容的提取
class Extractor(object):
    def __init__(self, file_path, rules):
        '''
        :param file_path:PDF file path
        :param rules: extract rules
        '''
        self.file_path = file_path
        self.rules = rules

    # 加载PDF文件
    def parse_pages(self):
        try:
            pages = []
            pdf = pdfplumber.open(self.file_path)
            print('parse file:{}   page num:{}'.format(os.path.basename(self.file_path), len(pdf.pages)))
            for index, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                if len(tables) < 1:
                    continue
                pages.append({'text': page.extract_text(), 'tables': tables, 'page': index + 1})
            return pages
        except Exception as e:
            print(e)
        return None

    # 提取特定类型表头的表格，规则有rules参数指定
    def extract_table_with_specific_header(self, pages):
        if pages is None or len(pages) < 1:
            # print('no-page...')
            return
        target_tables = []
        # 遍历所有页面
        for index, page in enumerate(pages):
            text = page['text']
            tables = page['tables']
            page_id = page['page'] - 1
            lines = re.split(r'\n+', text)
            # 遍历当前页面的所有行
            for ind, line in enumerate(lines):
                # 判定表头符合规则的表格
                if sum([1 for rule in self.rules['in-header'] if re.search(rule, line)]) > 0 and \
                        sum([1 for rule in self.rules['not-in-header'] if re.search(rule, line)]) == 0:
                    if ind >= len(lines) - 1:
                        break
                    cnt = ind + 1
                    next = lines[ind + 1]
                    if ind < len(lines) - 2 and re.search(r'单位[：:]', next):
                        next = lines[ind + 2]
                        cnt += 1
                    for ti, table in enumerate(tables):
                        if not table:
                            continue
                        first = ''.join([word for word in table[0] if word is not None])
                        # 表格是完整的情况
                        if first == re.sub(r'\s+', '', next):
                            tables[ti] = False
                            if index + 1 < len(pages) and len(pages[index + 1]['tables']) > 0:
                                table_next = pages[index + 1]['tables'][0]
                                fi = ''.join([item for item in table_next[0] if item is not None])
                                if re.sub(r'\s+', '', fi) in re.sub(r'\s+', '', pages[index + 1]['text'][1]):
                                    table += table_next
                            target_tables.append({'page': page_id + 1, 'method': 'exact', 'table': table,
                                                  'table-id': str(page_id + 1) + str(ti + 1)})
                        # 表格可能不完整的情况
                        elif first in re.sub(r'\s+', '', next):
                            tables[ti] = False
                            if index + 1 < len(pages) and len(pages[index + 1]['tables']) > 0:
                                table_next = pages[index + 1]['tables'][0]
                                fi = ''.join([item for item in table_next[0] if item is not None])
                                if re.sub(r'\s+', '', fi) in re.sub(r'\s+', '', pages[index + 1]['text'][1]):
                                    table += table_next
                            target_tables.append({'page': page_id + 1, 'method': 'guess', 'table': table,
                                                  'table-id': str(page_id + 1) + '-' + str(ti + 1)})
        return target_tables

    # 提取表格中存在指定类型信息的表格，规则由参数rules指定
    def extract_table_with_specific_info(self, pages):
        if pages is None or len(pages) < 1:
            # print('no-page...')
            return
        target_tables = []
        for index, page in enumerate(pages):
            tables = page['tables']
            page_id = page['page'] - 1
            for ti, table in enumerate(tables):
                st = str(table)
                if len(self.rules['in-table']) == sum([1 for in_tab in self.rules['in-table'] if in_tab in st]):
                    if len(self.rules['not-in-table']) == sum(
                            [1 for not_tab in self.rules['not-in-table'] if not not_tab in st]):
                        target_tables.append({'page': page_id + 1, 'method': 'content-in-table', 'table': table,
                                              'table-id': str(page_id + 1) + str(ti + 1)})
        return target_tables

    # 提取存在指定关键词的页面，关键词有rules指定
    def extract_specific_page(self, pages):
        if pages is None or len(pages) < 1:
            # print('no-page...')
            return
        target_pages = []
        for index, page in enumerate(pages):
            text = page['text']
            page_id = page['page']
            if len(self.rules['in-page']) == sum([1 for rule in self.rules['in-page'] if re.search(rule, text)]):
                target_pages.append({'page': page_id, 'text': text})
        return target_pages

    # 执行以上所有过程，返回提取结果
    def run(self):
        pages = self.parse_pages()
        if pages is None or len(pages) < 1:
            print('parse pdf error:', os.path.basename(self.file_path))
            return
        try:
            target_1 = self.extract_table_with_specific_header(pages)
            target_2 = self.extract_table_with_specific_info(pages)
            target_3 = self.extract_specific_page(pages)
            tables = []
            s = []
            for table in target_1:
                if table['table-id'] not in s:
                    s.append(table['table-id'])
                    tables.append(table)
            for table in target_2:
                if table['table-id'] not in s:
                    s.append(table['table-id'])
                    tables.append(table)
            return tables, target_3
        except Exception as e:
            print(e)


# 该类用来加载Excel，遍历地址获取PDF文件路径及缓存结果
class Util():
    def __init__(self, folder, out, demo):
        self.folder = folder
        self.out = out
        self.demo = demo

    # 加载Demo文件，获取rules
    def load_demo(self):
        print('load demo Excel:', os.path.basename(self.demo))
        book = xlrd.open_workbook(self.demo)
        sheet = book.sheet_by_index(0)
        in_header = sheet.col_values(0, 2, sheet.nrows)
        not_in_header = sheet.col_values(1, 2, sheet.nrows)
        in_table = sheet.col_values(2, 2, sheet.nrows)
        not_in_table = sheet.col_values(3, 2, sheet.nrows)
        in_page = sheet.col_values(4, 2, sheet.nrows)
        rules = {'in-header': in_header, 'not-in-header': not_in_header, 'in-table': in_table,
                 'not-in-table': not_in_table, 'in-page': in_page}
        for k, v in rules.items():
            rules[k] = [i for i in v if not re.sub(r'\s+', '', i) == '']
        return rules

    # 加载PDF文件，采用迭代遍历
    def load_folder(self):
        print('load folder:', self.folder)
        paths = []
        for dirpath, dirnames, filenames in os.walk(self.folder):
            for file in filenames:
                path = os.path.join(dirpath, file)
                if os.path.isfile(path) and (
                        os.path.splitext(path)[1] == '.pdf' or os.path.splitext(path)[1] == '.PDF'):
                    paths.append(path)
        return paths

    # 缓存结果
    def save_tmp(self, info, name, code, year):
        print('save tmp file:', name)
        if not os.path.isdir('tmp'):
            os.mkdir('tmp')
        tables = info[0]
        pages = info[1]

        # Excel样式
        style = xlwt.XFStyle()
        # background color
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 5
        style.pattern = pattern
        # border
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THICK
        borders.right = xlwt.Borders.THICK
        borders.top = xlwt.Borders.THICK
        borders.bottom = xlwt.Borders.THICK
        style.borders = borders
        # font
        font = xlwt.Font()
        font.name = 'Times New Roman'
        font.bold = True
        font.underline = False
        font.italic = False
        style.font = font

        # sheet style-2
        style2 = xlwt.XFStyle()
        # background color
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = 22
        style2.pattern = pattern
        # border
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THICK
        borders.right = xlwt.Borders.THICK
        borders.top = xlwt.Borders.THICK
        borders.bottom = xlwt.Borders.THICK
        style2.borders = borders
        # font
        font = xlwt.Font()
        font.name = 'Times New Roman'
        font.bold = True
        font.underline = False
        font.italic = False
        style2.font = font

        # sheet style-3
        style3 = xlwt.XFStyle()
        # background color
        pattern2 = xlwt.Pattern()
        pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern2.pattern_fore_colour = 3
        style3.pattern = pattern2
        # border
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THICK
        borders.right = xlwt.Borders.THICK
        borders.top = xlwt.Borders.THICK
        borders.bottom = xlwt.Borders.THICK
        style3.borders = borders
        # font
        font = xlwt.Font()
        font.name = 'Times New Roman'
        font.bold = True
        font.underline = False
        font.italic = False
        style3.font = font

        # 将数据写如Excel
        book = xlwt.Workbook()
        sheet1 = book.add_sheet('tables')
        sheet2 = book.add_sheet('pages')

        for ind, page in enumerate(pages):
            page_num = page['page']
            text = page['text']
            sheet2.write(ind, 0, name)
            sheet2.write(ind, 1, code)
            sheet2.write(ind, 2, year)
            sheet2.write(ind, 3, page_num)
            sheet2.write(ind, 4, 'search page')
            sheet2.write(ind, 5, text, style)

        # save table
        i = 0
        for ti, table in enumerate(tables):
            page = table['page']
            method = table['method']
            table_content = table['table']
            if method == 'exact':
                sty = style
            elif method == 'guess':
                sty = style2
            else:
                sty = style3
            for index, row in enumerate(table_content):
                sheet1.write(i, 0, name)
                sheet1.write(i, 1, code)
                sheet1.write(i, 2, year)
                sheet1.write(i, 3, page)
                sheet1.write(i, 4, method)
                for ind, one in enumerate(row):
                    sheet1.write(i, 5 + ind, one if one is not None else '', sty)
                i += 1
            i += 1

        book.save('tmp\\' + name + '.tmp.xls')


# 单个文件运行的完整流程，从加载文件到缓存结果的全过程，如果只想使用单线程运行程序，则在主函数中调用该函数即可
def run(rules, file, util):
    extractor = Extractor(file, rules)
    info = extractor.run()
    code = re.findall(r'\d{6}', os.path.basename(file))[0]
    year = re.findall(r'\d{8}', os.path.basename(file))[0]
    if info is None or len(info[0]) < 1:
        with open('noResult.txt', 'a', encoding='utf-8') as fp:
            fp.write(file + '\n')
    else:
        util.save_tmp(info, os.path.basename(file), code, year)


# -------------------
# 以下两个函数是为了加快执行速度而启用的多线程+多进程模式，计算密集型任务状态下进程越多越好（不多于机器CPU核心数）
# -----------------
# 多线程：每次会启动跟files数量相对应的线程来执行，但只能执行在一个CPU核心中
# multiple threads
def batch_processor(func, rules, files, util):
    thread_pool = []
    for index, file in enumerate(files):
        th = threading.Thread(target=func, args=(rules, file, util))
        # print('running thread:', th.name)
        th.start()
        thread_pool.append(th)
    for th in thread_pool:
        # print('waiting for thread:', th.name)
        th.join()


# 多进程：启动4个进程执行，每个进程中运行多线程，CPU有几个核心就使用几个进程，一般机器多为双核心四进程，此时4进程可占满CPU运行，效能最大
# multiple processors
def multi_processor_run(func, sub_func, files, rules, util):
    pool = multiprocessing.Pool(processes=4)
    cnt = 0
    batch_size = 5
    while cnt < len(files):
        rear = cnt + batch_size
        if rear > len(files):
            rear = len(files)
        batch = files[cnt + 0:rear]
        pool.apply_async(func, (sub_func, rules, batch, util))
        cnt += batch_size
    pool.close()
    pool.join()


# 该函数将缓存在本地目录tmp文件夹下的所有临时Excel文件结果整合到一个Excel中
# re-format result
def re_format(sheet_size):
    print('re-format file...')
    files = os.listdir('tmp')
    paths = []
    new_book = xlwt.Workbook()
    for file in files:
        if os.path.isfile(os.path.join('tmp', file)) and '.tmp.xls' in file:
            paths.append(os.path.join('tmp', file))

    style = xlwt.XFStyle()
    # background color
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 5
    style.pattern = pattern
    # border
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THICK
    borders.right = xlwt.Borders.THICK
    borders.top = xlwt.Borders.THICK
    borders.bottom = xlwt.Borders.THICK
    style.borders = borders
    # font
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.bold = True
    font.underline = False
    font.italic = False
    style.font = font

    # sheet style-2
    style2 = xlwt.XFStyle()
    # background color
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = 22
    style2.pattern = pattern
    # border
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THICK
    borders.right = xlwt.Borders.THICK
    borders.top = xlwt.Borders.THICK
    borders.bottom = xlwt.Borders.THICK
    style2.borders = borders
    # font
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.bold = True
    font.underline = False
    font.italic = False
    style2.font = font

    # sheet style-3
    style3 = xlwt.XFStyle()
    # background color
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = 3
    style3.pattern = pattern2
    # border
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THICK
    borders.right = xlwt.Borders.THICK
    borders.top = xlwt.Borders.THICK
    borders.bottom = xlwt.Borders.THICK
    style3.borders = borders
    # font
    font = xlwt.Font()
    font.name = 'Times New Roman'
    font.bold = True
    font.underline = False
    font.italic = False
    style3.font = font

    tab_cnt = 1
    page_cnt = 1
    tab_rows = 0
    page_rows = 0
    sheet2 = new_book.add_sheet('pages-' + str(0))
    sheet1 = new_book.add_sheet('tables-' + str(0))
    sheet1.write(0, 0, 'File')
    sheet1.write(0, 1, 'Code')
    sheet1.write(0, 2, 'Date')
    sheet1.write(0, 3, 'Page')
    sheet1.write(0, 4, 'Method')
    sheet2.write(0, 0, 'File')
    sheet2.write(0, 1, 'Code')
    sheet2.write(0, 2, 'Date')
    sheet2.write(0, 3, 'Page')
    sheet2.write(0, 4, 'Method')

    for index, file in enumerate(paths):
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_name('tables')
        sheet_pages = book.sheet_by_name('pages')
        tab_rows += sheet.nrows
        page_rows += sheet_pages.nrows

        for row in range(sheet.nrows):
            if len(sheet.row_values(row)) < 5:
                sty1 = None
            elif sheet.row_values(row)[4] == 'exact':
                sty1 = style
            elif sheet.row_values(row)[4] == 'guess':
                sty1 = style2
            elif sheet.row_values(row)[4] == 'content-in-table':
                sty1 = style3
            else:
                sty1 = None
            for col, val in enumerate(sheet.row_values(row)):
                if col > 4:
                    if sty1 is not None:
                        sheet1.write(tab_cnt, col, val, sty1)
                    else:
                        sheet1.write(tab_cnt, col, val)
                else:
                    sheet1.write(tab_cnt, col, val)
            tab_cnt += 1
        tab_cnt += 1
        for row in range(sheet_pages.nrows):
            if len(sheet_pages.row_values(row)) < 5:
                sty2 = None
            elif sheet_pages.row_values(row)[4] == 'exact':
                sty2 = style
            elif sheet_pages.row_values(row)[4] == 'guess':
                sty2 = style2
            elif sheet_pages.row_values(row)[4] == 'content-in-table':
                sty2 = style3
            else:
                sty2 = None
            for col, val in enumerate(sheet_pages.row_values(row)):
                if col > 4:
                    if sty2 is not None:
                        sheet2.write(page_cnt, col, val, sty2)
                    else:
                        sheet2.write(page_cnt, col, val)
                else:
                    sheet2.write(page_cnt, col, val)
            page_cnt += 1
        page_cnt += 1
        if tab_rows >= sheet_size:
            tab_rows = 0
            tab_cnt = 1
            sheet1 = new_book.add_sheet('tables-' + str(index))
            sheet1.write(0, 0, 'File')
            sheet1.write(0, 1, 'Code')
            sheet1.write(0, 2, 'Date')
            sheet1.write(0, 3, 'Page')
            sheet1.write(0, 4, 'Method')

        if page_rows >= sheet_size:
            page_rows = 0
            page_cnt = 1
            sheet2 = new_book.add_sheet('pages-' + str(index))
            sheet2.write(0, 0, 'File')
            sheet2.write(0, 1, 'Code')
            sheet2.write(0, 2, 'Date')
            sheet2.write(0, 3, 'Page')
            sheet2.write(0, 4, 'Method')

    new_book.save('tables.xls')


# 程序执行入口：主函数
if __name__ == '__main__':
    # 此命令是为在Windows环境下打包exe时正确引入多进程模块而添加的，在Python解释器中运行代码这一行是不必要的，当然添加之后也无妨
    multiprocessing.freeze_support()
    # 程序运行需要的参数
    # paras
    base_dir = r'./'  # 程序工作目录设定为本程序所在的目录
    out_path = base_dir + r'\result.xls'  # 输出结果文件名称
    demo = base_dir + r'\Demo.xlsx'  # Demo文件名称

    # 新建noResult.txt文件，用来保存没有结果的PDF文件名称
    with open('noResult.txt', 'w', encoding='utf-8')as fp:
        fp.write(str(datetime.now()) + '\n')

    # 初始化Util类
    util = Util(base_dir + '\\test', out_path, demo)
    rules = util.load_demo()
    folder = util.load_folder()

    # 执行多进程，，但仅执行单线程模式时这里可替换为run函数
    multi_processor_run(batch_processor, run, folder, rules, util)

    # 保存结果：5000代表每个Excel的单个sheet最多5000行，超过则会新建sheet
    # save...
    re_format(5000)
    # 移除临时文件，这些临时文件在程序运行过程中会保存在当前目录的tmp文件夹内，其中每个Excel文件保存的是单个PDF文件的结果，最终这些结果将会通过re_format函数整合到一个Excel中，当想要保留这些结果时，可将下面一行代码注释掉
    shutil.rmtree('tmp')

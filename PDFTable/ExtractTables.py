# -*- coding:utf-8 -*-
import pdfplumber
import re
import xlrd
from xlutils.copy import copy
import os
import xlwt
import threading
import multiprocessing

from . import multi_process


class Extractor(object):
    def __init__(self, file_path):
        self.file_path = file_path

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

    def extract_tables(self, pages):
        if pages is None or len(pages) < 1:
            # print('no-page...')
            return
        target_tables = []
        tab_cnt = 0
        for index, page in enumerate(pages):
            text = page['text']
            tables = page['tables']
            page_id = page['page'] - 1
            lines = re.split(r'\n+', text)
            for ind, table in enumerate(tables):
                first = ''.join([item if item is not None else '' for item in table[0]])
                flag = False
                for i, line in enumerate(lines):
                    if re.search(r'长期借款|长期负债', line):
                        flag = True
                    if not flag:
                        continue
                    if i == 1 and ind > 0 and tab_cnt > 0 and re.sub(r'\s+', '', line) == first:
                        title = lines[i - 1]
                        if re.search(r'单位[:：\s]+元', title):
                            title = lines[i - 2]
                        target_tables.append(
                            {'page': page_id, 'table-cnt': ind + 1, 'type': 'merged', 'reliability': 'exact',
                             'title': title, 'table': target_tables[tab_cnt - 1]['table'] + table})
                        tab_cnt += 1
                    elif i == 1 and ind > 0 and tab_cnt > 0 and first in re.sub(r'\s+', '', line):
                        title = lines[i - 1]
                        if re.search(r'单位[:：\s]+元', title):
                            title = lines[i - 2]
                        target_tables.append(
                            {'page': page_id, 'table-cnt': ind + 1, 'type': 'merged', 'reliability': 'part',
                             'title': title, 'table': target_tables[tab_cnt - 1]['table'] + table})
                        tab_cnt += 1
                    elif i > 1 and re.sub(r'\s+', '', line) == first:
                        title = lines[i - 1]
                        if re.search(r'单位[:：\s]+元', title):
                            title = lines[i - 2]
                        target_tables.append(
                            {'page': page_id, 'table-cnt': ind + 1, 'type': 'origin', 'reliability': 'exact',
                             'title': title, 'table': table})
                        tab_cnt += 1
                    elif i > 1 and first in re.sub(r'\s+', '', line):
                        title = lines[i - 1]
                        if re.search(r'单位[:：\s]+元', title):
                            title = lines[i - 2]
                        target_tables.append(
                            {'page': page_id, 'table-cnt': ind + 1, 'type': 'origin', 'reliability': 'part',
                             'title': title, 'table': table})
                        tab_cnt += 1
        return target_tables

    def run(self):
        pages = self.parse_pages()
        tables = self.extract_tables(pages)
        return tables


def load_folder(folder):
    files = os.listdir(folder)
    paths = []
    for file in files:
        path = os.path.join(folder, file)
        if os.path.isfile(path) and os.path.splitext(path)[1] == '.pdf':
            paths.append(path)
    return paths


def add_sheets(files, out_path):
    # add sheets
    print('add sheets...')
    book = xlwt.Workbook()
    for index, file in enumerate(files):
        name = re.sub(r'[\[\]():：]+', '', os.path.basename(file))
        sheet_name = '{:0>4d}-'.format(index + 1) + name[:20] + '...' if len(name) > 20 else name
        book.add_sheet(sheet_name, cell_overwrite_ok=True)
    book.save(out_path)


def saver(tables, out_path, sheet_name):
    if tables is None or len(tables) < 1:
        print('no table to save...')
        return
    print('run saver...')
    # sheet style
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

    # sheet style-3
    style3 = xlwt.XFStyle()
    # background color
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2.pattern_fore_colour = 3
    style3.pattern = pattern2

    book = xlrd.open_workbook(out_path, formatting_info=True)
    copy_book = copy(book)
    sheet_copy = copy_book.get_sheet(sheet_name)
    rows = 0
    for ind, table in enumerate(tables):
        page = table['page']
        table_cnt = table['table-cnt']
        type = table['type']
        reliability = table['reliability']
        title = table['title']
        table = table['table']

        style4 = style3 if reliability == 'exact' else style2

        # page num
        sheet_copy.write(rows + ind, 0, '页码', style4)
        sheet_copy.write(rows + ind + 1, 0, page, style4)
        # table num
        sheet_copy.write(rows + ind, 1, '表格编号', style4)
        sheet_copy.write(rows + ind + 1, 1, table_cnt, style4)
        # title
        sheet_copy.write(rows + ind, 2, '表头', style4)
        sheet_copy.write(rows + ind + 1, 2, title if title is not None else '未知', style4)
        # type
        sheet_copy.write(rows + ind, 3, '类型', style4)
        sheet_copy.write(rows + ind + 1, 3, type, style4)
        # reliability
        sheet_copy.write(rows + ind, 4, '可靠性', style4)
        sheet_copy.write(rows + ind + 1, 4, '精确匹配' if reliability == 'exact' else '不完整匹配', style4)
        # table
        sheet_copy.write(rows + ind + 2, 0, '表-{} 内容'.format((ind + 1)))
        for row_cnt, row in enumerate(table):
            for col_cnt, col in enumerate(row):
                sheet_copy.write(rows + ind + 3 + row_cnt, col_cnt, col, style)
        rows += 3 + len(table)
    copy_book.save(out_path)


def batch_parser(files, out_path):
    for index, file in enumerate(files):
        print('progress:', index + 1, '/', len(files))
        extractor = Extractor(file)
        tables = extractor.run()
        name = re.sub(r'[\[\]()：:]+', '', os.path.basename(file))
        sheet_name = '{:0>4d}-'.format(index + 1) + name[:20] + '...' if len(name) > 20 else name
        saver(tables, out_path, sheet_name)


def run(file, index, out_path):
    extractor = Extractor(file)
    tables = extractor.run()
    name = re.sub(r'[\[\]()：:]+', '', os.path.basename(file))
    sheet_name = '{:0>4d}-'.format(index + 1) + name[:20] + '...' if len(name) > 20 else name
    lock = threading.Lock()
    lock.acquire()
    saver(tables, out_path, sheet_name)
    lock.release()


# multiple threads
def batch_processor(batch, batch_size, files, out_path):
    thread_pool = []
    for index, file in enumerate(files):
        th = threading.Thread(target=run, args=(file, batch * batch_size + index, out_path))
        # print('running thread:', th.name)
        th.start()
        thread_pool.append(th)
    for th in thread_pool:
        # print('waiting for thread:', th.name)
        th.join()


# multiple processors
def multi_processor_run(func, files, out_path):
    pool = multiprocessing.Pool(processes=4)
    cnt = 0
    batch_size = 5
    while cnt < len(files):
        rear = cnt + batch_size
        if rear > len(files):
            rear = len(files)
        batch = files[cnt + 0:rear]
        pool.apply_async(func, (int(cnt / batch_size), batch_size, batch, out_path))
        cnt += batch_size
    pool.close()
    pool.join()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    # paras
    base_dir = r'C:\Users\fanyu\Desktop\Project\PaidDevelopment\11_Table_PDF2Excel_￥500\data'
    out_path = r'result.xls'
    files = load_folder(base_dir)
    add_sheets(files, out_path)
    # run...
    # batch_parser(files, out_path)
    multi_processor_run(batch_processor, files, out_path)

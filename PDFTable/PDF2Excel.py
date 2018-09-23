# -*- coding:utf-8 -*-
import pdfplumber
import re
import xlrd
from xlutils.copy import copy
import os
import xlwt


class Extractor(object):
    def __init__(self, file_path, rules):
        self.file_path = file_path
        self.rules = rules

    def parse_pages(self):
        try:
            pages = []
            pdf = pdfplumber.open(self.file_path)
            print('parse file:{}   page num:{}'.format(os.path.basename(self.file_path), len(pdf.pages)))
            for index, page in enumerate(pdf.pages):
                pages.append({'text': page.extract_text(), 'tables': page.extract_tables(), 'page': index + 1})
            return pages
        except Exception as e:
            print(e)
        return None

    def extract_tables(self, pages):
        if pages is None or len(pages) < 1:
            # print('no-page...')
            return
        target_tables = []
        # find sections
        file_name = os.path.basename(self.file_path)
        # search section
        page_num = len(pages)
        sections = []
        section_cnt = 0
        # print('search sections...')
        for index, page in enumerate(pages):
            lines = re.split(r'\n+', page['text'])
            for ind, text in enumerate(lines):
                if re.search(r'^\d+[、.][\u4e00-\u9fa5\s]+$', text):
                    sections.append({'name': text, 'page-st': index + 1, 'page-ed': page_num, 'file': file_name})
                    section_cnt += 1
                    if section_cnt > 1:
                        sections[section_cnt - 2]['page-ed'] = sections[section_cnt - 1]['page-st']
        # search table
        # print('search tables in sections(%s)...' % len(sections))
        for section in sections:
            skip = False
            for word in self.rules['title-not-in-word-or']:
                if word in section['name']:
                    skip = True
                    break
            if skip:
                continue
            for word in self.rules['title-in-word-or']:
                if word in section['name']:
                    for i, page in enumerate(pages[section['page-st'] - 1: section['page-ed']]):
                        for ii, table in enumerate(page['tables']):
                            target_tables.append(
                                {'page': page['page'], 'table-cnt': ii + 1, 'method': 'in-section',
                                 'section-name': section['name'], 'type': None, 'reliability': None, 'header': None,
                                 'table': table})
                    break
        # search table
        # print('search tables with "table-inner-in-word-and" rule...')
        used_pages = []
        for section in sections:
            skip = False
            for word in self.rules['title-not-in-word-or']:
                if word in section['name']:
                    skip = True
                    break
            if skip:
                continue
            for word in self.rules['title-in-word-or']:
                if word in section['name']:
                    used_pages += range(section['page-st'] - 1, section['page-ed'])

        for index, page in enumerate(pages):
            if index in used_pages:
                continue
            for ii, table in enumerate(page['tables']):
                skip = False
                word_flag = [False for _ in self.rules['table-inner-in-word-and']]
                cnt = 0
                for iii, row in enumerate(table):
                    for word in self.rules['table-inner-not-in-word-or']:
                        if word in row:
                            skip = True
                            break
                    if skip:
                        break
                    for ind, word in enumerate(self.rules['table-inner-in-word-and']):
                        if cnt == 0 and word in row:
                            word_flag[cnt] = True
                            cnt += 1
                        elif ind > cnt and word in row:
                            word_flag[cnt] = True
                            cnt += 1
                if False not in word_flag:
                    target_tables.append(
                        {'page': index + 1, 'table-cnt': ii + 1, 'method': 'in-table', 'section-name': None,
                         'type': None, 'reliability': None, 'header': None, 'table': table})
        return target_tables

    def find_header(self, pages):
        tables = self.extract_tables(pages)
        if tables is None or len(tables) < 1:
            # print('no table...')
            return
        for index, page in enumerate(pages):
            text = page['text']
            for tab_cnt, table in enumerate(tables):
                if not tables[tab_cnt]['page'] == index + 1:
                    continue
                row = [item if item is not None else '' for item in tables[tab_cnt]['table'][0]]
                first_row = ''.join(row)
                lines = re.split(r'\n+', text)
                for ind, line in enumerate(lines):
                    if ind == 1 and tab_cnt > 0 and re.sub(r'\s+', '', line) == first_row:
                        tables[tab_cnt - 1]['table'] += tables[tab_cnt]['table']
                        tables[tab_cnt - 1]['type'] = 'merged'
                        tables[tab_cnt - 1]['reliability'] = 'exact'
                        tables[tab_cnt]['type'] = 'discarded'
                        break
                    elif ind == 1 and tab_cnt > 0 and first_row in re.sub(r'\s+', '', line):
                        tables[tab_cnt - 1]['table'] += tables[tab_cnt]['table']
                        tables[tab_cnt - 1]['type'] = 'merged'
                        tables[tab_cnt - 1]['reliability'] = 'guess'
                        tables[tab_cnt]['type'] = 'discarded'
                        break
                    elif ind > 1 and first_row == re.sub(r'\s+', '', line):
                        tables[tab_cnt]['type'] = 'origin'
                        tables[tab_cnt]['reliability'] = 'exact'
                        header = lines[ind - 1]
                        if re.search(r'单位[:：\s]+元', header):
                            header = lines[ind - 2]
                        tables[tab_cnt]['header'] = header
                        break
                    elif ind > 1 and first_row in re.sub(r'\s+', '', line):
                        tables[tab_cnt]['type'] = 'origin'
                        tables[tab_cnt]['reliability'] = 'guess'
                        header = lines[ind - 1]
                        if re.search(r'单位[:：\s]+元', header):
                            header = lines[ind - 2]
                        tables[tab_cnt]['header'] = header
                        break
        return tables

    def run(self):
        pages = self.parse_pages()
        tables = self.find_header(pages)
        return tables


def load_folder(folder):
    files = os.listdir(folder)
    paths = []
    for file in files:
        path = os.path.join(folder, file)
        if os.path.isfile(path) and os.path.splitext(path)[1] == '.pdf':
            paths.append(path)
    return paths


def rules():
    rule = {'title-in-word-or': ['长期借款', '长期负债'], 'title-not-in-word-or': ['应收款', '短期', '担保', '到期'],
            'table-inner-in-word-and': ['贷款单位', '利率', '起始日'], 'table-inner-not-in-word-or': ['关联']}
    return rule


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
        if table['type'] == 'discarded':
            continue
        page = table['page']
        table_cnt = table['table-cnt']
        method = table['method']
        section_name = table['section-name']
        type = table['type']
        reliability = table['reliability']
        header = table['header']
        table = table['table']
        style4 = style3 if reliability == 'exact' else style2
        # page num
        sheet_copy.write(rows + ind, 0, '页码', style4)
        sheet_copy.write(rows + ind + 1, 0, page, style4)
        # table num
        sheet_copy.write(rows + ind, 1, '表格编号', style4)
        sheet_copy.write(rows + ind + 1, 1, table_cnt, style4)
        # search method
        sheet_copy.write(rows + ind, 2, '搜寻方式', style4)
        sheet_copy.write(rows + ind + 1, 2, method, style4)
        # section name
        sheet_copy.write(rows + ind, 3, '从属于类', style4)
        sheet_copy.write(rows + ind + 1, 3, section_name if section_name is not None else '无', style4)
        # header
        sheet_copy.write(rows + ind, 4, '表头', style4)
        sheet_copy.write(rows + ind + 1, 4, header if header is not None else '未知', style4)
        # type
        if type == 'merged':
            TYPE = '跨页合并'
        elif type == 'discarded':
            TYPE = '已被合并'
        elif type == 'origin':
            TYPE = '原始表格'
        else:
            TYPE = '未知'
        sheet_copy.write(rows + ind, 5, '类型', style4)
        sheet_copy.write(rows + ind + 1, 5, TYPE, style4)
        # reliability
        sheet_copy.write(rows + ind, 6, '可靠性', style4)
        sheet_copy.write(rows + ind + 1, 6, '精确匹配' if reliability == 'exact' else '猜测匹配', style4)

        # table
        sheet_copy.write(rows + ind + 2, 0, '表-{} 内容'.format((ind + 1)))
        for row_cnt, row in enumerate(table):
            for col_cnt, col in enumerate(row):
                sheet_copy.write(rows + ind + 3 + row_cnt, col_cnt, col, style)
        rows += 3 + len(table)
    copy_book.save(out_path)


def batch_parser(files, rules, out_path):
    add_sheets(files, out_path)
    for index, file in enumerate(files):
        print('progress:', index + 1, '-', len(files))
        extractor = Extractor(file, rules)
        tables = extractor.run()
        name = re.sub(r'[\[\]()：:]+', '', os.path.basename(file))
        sheet_name = '{:0>4d}-'.format(index + 1) + name[:20] + '...' if len(name) > 20 else name
        saver(tables, out_path, sheet_name)


if __name__ == '__main__':
    # paras
    base_dir = r'C:\Users\fanyu\Desktop\Project\PaidDevelopment\11_Table_PDF2Excel_￥500\data'
    out_path = r'result.xls'
    rule = rules()
    files = load_folder(base_dir)

    # run...
    batch_parser(files, rule, out_path)

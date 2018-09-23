import pdfplumber
import re
import os
import xlrd
from xlutils.copy import copy
import threading
import time
import multiprocessing
from datetime import datetime


def search_page(path):
    print('load file:', os.path.basename(path))
    pdf = pdfplumber.open(path)
    pages = pdf.pages
    print('total page:', len(pages))
    print('searching pdf...')
    target = []
    st_flag = False
    for ind, page in enumerate(pages):
        # print('parse page:', ind + 1)
        text = page.extract_text()
        lines = re.split(r'\n+', text)
        for index, line in enumerate(lines):
            if not st_flag and re.match(r'\s*\d+[、.\s]+税金及附加\s*$', line):
                st_flag = True
                continue
            if st_flag and '合计' not in line:
                target.append(line)
            elif st_flag and '合计' in line:
                return target


def target(lines, rules):
    target_info = []
    for line in lines:
        items = line.split(' ')
        if not isinstance(items, list) or len(items) < 2:
            continue
        for rule in rules:
            if rule == items[0]:
                target_info.append(items)
    return target_info


def saver(out_path, rules):
    print('save to file:', os.path.basename(out_path))
    files = os.listdir('./')

    book = xlrd.open_workbook(out_path)
    cbook = copy(book)
    sheet = book.sheet_by_index(0)
    csheet = cbook.get_sheet('Sheet1')

    for index, file in enumerate(files):
        info = []
        if os.path.isfile(file) and os.path.splitext(file)[1] == '.tmp':
            code = re.split('-', file)[1]
            name = re.split(r'[\-:：]', file)[2]
            with open(file, 'r', encoding='utf-8') as fp:
                lines = fp.readlines()
            for line in lines:
                info.append(line.split('|'))
            csheet.write(sheet.nrows + index, 0, str(code))
            csheet.write(sheet.nrows + index, 1, str(name))
            for ind, one in enumerate(info):
                if len(one) < 2:
                    continue
                for ind2, rule in enumerate(rules):
                    if one[0] == rule:
                        csheet.write(sheet.nrows + index, ind2 + 2, str(one[1]))
    cbook.save(out_path)


def load_demo(path):
    print('load demo Excel:', os.path.basename(path))
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    row = sheet.row_values(1, 0, sheet.ncols)
    return row


def load_folder(folder_path):
    paths = []
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for file in filenames:
            path = os.path.join(dirpath, file)
            if os.path.isfile(path) and (os.path.splitext(path)[1] == '.pdf' or os.path.splitext(path)[1] == '.PDF'):
                paths.append(path)
    return paths


def run(path, row):
    # print('processing on file:', os.path.basename(path))
    pdf = os.path.basename(path)
    rule = row[2:]
    try:
        lines = search_page(path)
    except Exception as e:
        print('parse pdf error:', e)
        with open('errorList.txt', 'a', encoding='utf-8') as fp:
            fp.write('parse pdf error:' + path + '\n')
        return
    try:
        info = target(lines, rule)
        if not isinstance(info, list) or len(info) < 1:
            print('find nothing in file:', pdf)
            with open('errorList.txt', 'a', encoding='utf-8') as fp:
                fp.write('find nothing in file:' + path + '\n')
            return
    except Exception as e:
        print('search lines error:', e)
        with open('errorList.txt', 'a', encoding='utf-8') as fp:
            fp.write('search lines error:' + path + '\n')
        return
    try:
        with open(pdf + '.tmp', 'w', encoding='utf-8') as fp:
            for one in info:
                for item in one:
                    fp.write(str(item) + '|')
                fp.write('\n')
    except Exception as e:
        print('save tmp file error:', e)


def multi_threads(paths, row):
    th_pool = []
    for i, path in enumerate(paths):
        th = threading.Thread(target=run, args=(path, row))
        th.start()
        th_pool.append(th)
    for th in th_pool:
        th.join()


def batch_parser(folder, demo):
    try:
        paths = load_folder(folder)
        row = load_demo(demo)
    except Exception as e:
        print('load Excel/folder error:', e)
        return
    print('total {} files.'.format(len(paths)))
    pool = multiprocessing.Pool(processes=4)
    for i in range(0, len(paths), 5):
        path = paths[i:(i + 5 if (i + 5 <= len(paths) - 1) else len(paths))]
        pool.apply_async(func=multi_threads, args=(path, row))
    pool.close()
    pool.join()


if __name__ == '__main__':
    # paras
    base_dir = r'C:\Users\fanyu\Desktop\PaidProject\16_Tables_PDF_extractor'
    demo = r'C:\Users\fanyu\Desktop\PaidProject\16_Tables_PDF_extractor\Demo.xls'
    with open('errorList.txt', 'w', encoding='utf-8')as fp:
        fp.write(str(datetime.now()) + '\n\n')
    # batch_parser(base_dir, demo)
    saver(demo, load_demo(demo)[2:])
    print('Program finished!')

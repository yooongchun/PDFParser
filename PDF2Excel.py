# -*- coding:utf-8 -*-
#

"""
Extract Authors name
"""

__author__ = "yooongchun"
__email__ = "yooongchun@foxmail.com"

import re
import os
import xlrd
from xlutils.copy import copy
import xlwt


class AuthorInfo2(object):
    def __init__(self, txt_path):
        if os.path.isfile(txt_path) and os.path.splitext(txt_path)[1] == '.txt':
            self.__path = txt_path
        else:
            print('输入路径错误，确保为txt格式的文件...')
            return

    def find_email(self, lines):
        emails = []
        for index, line in enumerate(lines):
            if '@' in line and re.search(r'[\[\]\-\w()._]+@[\w\-_]+\.+[\w\-_]+', line):
                items = re.findall(r'[\[\]\-\w()._]+@[\w\-_.]+', line)
                items = [re.sub(r'[\[\]()]+', '', item) for item in items]
                for i, item in enumerate(items):
                    if re.match(r'^\d[a-zA-Z()\[\]\-_.]+@.+', item):
                        items[i] = re.sub(r'^\d', '', item)
                    if re.match(r'.+@.+\d$', item):
                        items[i] = re.sub(r'\d$', '', item)
                    if '@' in str(lines[index - 5:index]) or '@' in str(lines[index + 1:index + 5]) or re.search(
                            r'abstract|keywords|introduction|correspondence',
                            str(lines[index - 20:index + 10]).lower()) or 'Corresponding' in line:
                        pass
                    else:
                        items.pop(i)
                emails += [{'index': index + 1, 'email': item} for item in items]
        return emails

    def match_author(self, emails, lines):
        author = []
        for index, _email in enumerate(emails):
            ind = _email['index']
            email = _email['email']
            base = email.split('@')[0]
            rear = email.split('@')[1]
            if re.match(r'^\d+', base) or 'qq' in rear:
                author.append({'name': '', 'email': email, 'method': 'skip'})
                continue
            # match author
            cnt = ind - 1
            while cnt > 0 and not re.match(r'^\s*\d+\s*\n$', lines[cnt]):
                cnt -= 1
            block = lines[cnt + 1:ind]
            block.reverse()
            if not block or len(block) < 1:
                continue
            find_flag = False
            for line in block:
                line = line.replace(email, '')
                if '@' in line or re.match(r'^\s+$|^\s*\n', line):
                    continue
                line = re.sub(r'\s+\n', ' ', line)
                items = re.split(r'By|by|BY|Author|author|\s+and\s+|,|:|\*|\d|\s{5,100}|&', line)
                names = [item for item in items if not re.sub(r'\s+', '', item) == '']
                names = [name for name in names if
                         len([n for n in re.split('\s+', name) if not re.sub(r'\s+', '', n) == '']) <= 4]
                scores = []
                for cc, name in enumerate(names):
                    if 'of' in name or '_' in name or 'universit' in name.lower() or 'learning' in name.lower():
                        continue
                    if 'for' in name:
                        continue
                    if re.search(r'[A-Z][a-z]$', name):
                        name = re.sub(r'[a-z]$', '', name)
                    name = re.sub(r'^\s*\d\s*|\s*\d\s*$', '', name)
                    name = re.sub(r'\s+', ' ', name)
                    name = re.sub(r'[()\[\]*;]', '', name)
                    name = re.sub(r'^[a-z]', '', name)
                    name = re.sub(r'^\s+|\s+$', '', name)
                    names[cc] = name

                    base = re.sub(r'\d+', '', base)
                    score = self.name_similarity(name, base)
                    scores.append(score)
                if len(scores) > 0 and max(scores) > 0:
                    find_flag = True
                    index_max = scores.index(max(scores))
                    author.append({'name': names[index_max], 'email': email, 'method': 'similarity:%d' % max(scores)})
                    break
            if not find_flag:
                author.append({'name': '', 'email': email, 'method': 'not-found'})

        author = sorted(author, key=lambda au: len(au['name']), reverse=True)
        tmp = [au for au in author if not au['name'] == '']
        tmp = sorted(tmp, key=lambda t: int(re.findall(r'\d+', t['method'])[0]), reverse=True)
        author[0:len(tmp)] = tmp

        return author

    def run(self):
        with open(self.__path, 'r', encoding='utf-8') as fp:
            lines = fp.readlines()
        author = self.match_author(self.find_email(lines), lines)
        return author

    def name_similarity(self, name_a, email_base):
        name_a = name_a.lower()
        name_b = email_base.lower()
        # full name match
        f1 = re.sub(r'\s+|\.', '', name_a)
        f2 = re.sub(r'[\-_.]', '', name_b)
        if f1 == f2:
            return 100
        # part name match
        items1 = re.split(r'\s+|\.', name_a)
        items2 = re.split(r'[\s+\-_.]', name_b)
        for s1 in items1:
            for s2 in items2:
                if s1 == s2 and len(s1) > 2:
                    return 50
        # max-sub string match
        ms1 = re.split(r'\s+|\.', name_a)
        for ms in ms1:
            if len(ms) > 2 and ms in name_b:
                return 25
        ms2 = re.split(r'\s+|\.|-|_', name_b)
        for ms in ms2:
            if len(ms) > 2 and ms in name_a:
                return 25
        # first-letter
        rule = ''
        for a in name_b:
            rule += a + '[a-z\-.]'
        if re.search(rule, name_a):
            return 15
        # over-ride rate
        C1 = re.sub(r'\s+|\.|-|_|\d+', '', name_a)
        C2 = re.sub(r'\s+|\.|-|_|\d+', '', name_b)
        if len(C1) < 3 or len(C2) < 3:
            return 0
        cnt_c2 = 0
        for c1 in C1:
            if cnt_c2 < len(C2) and c1 == C2[cnt_c2]:
                cnt_c2 += 1
        if cnt_c2 > 4:
            return 10
        # cnt_c1 = 0
        # for c2 in C2:
        #     if cnt_c1 < len(C1) and c2 == C1[cnt_c1]:
        #         cnt_c1 += 1
        # if cnt_c1 > 4:
        #     return 12.5
        return 0


def add_sheet(excel_path, names):
    key_words = ['Name', 'Email', 'Matching Method']
    book = xlwt.Workbook()
    for i, name in enumerate(names):
        if len(name) > 20:
            name = name[:20] + '...'

        name = '{:0>4d}_{}'.format(i, name)
        sheet = book.add_sheet(name, cell_overwrite_ok=True)
        for i in range(len(key_words)):
            sheet.write(4, i, key_words[i])
    book.save(excel_path)


def handle_folder(folder_path, excel_path):
    files = os.listdir(folder_path)
    txt_paths = []
    for file in files:
        path = os.path.join(folder_path, file)
        if os.path.isfile(path) and os.path.splitext(file)[1] == '.txt':
            txt_paths.append(file)
    add_sheet(excel_path, txt_paths)


def save2Excel(excel_path, sheet_name, author):
    book = xlrd.open_workbook(excel_path)
    copy_book = copy(book)
    sheet_copy = copy_book.get_sheet(sheet_name)
    cnt_name = 0
    for index, info in enumerate(author):
        value = info['name']
        sheet_copy.write(index + 5, 0, value)
        if not value == '':
            cnt_name += 1
        value = info['email']
        sheet_copy.write(index + 5, 1, value)
        value = info['method']
        sheet_copy.write(index + 5, 2, value)
    sheet_copy.write(0, 0, 'Total Email')
    sheet_copy.write(0, 1, '%d' % len(author))
    sheet_copy.write(1, 0, 'Total Author')
    sheet_copy.write(1, 1, '%d' % cnt_name)
    sheet_copy.write(2, 0, 'Extracted Rate')
    sheet_copy.write(2, 1, '%.2f' % (cnt_name / len(author)) if len(author) > 0 else 0)

    copy_book.save(excel_path)


if __name__ == "__main__":
    # paras
    base_dir = r'C:\Users\fanyu\Desktop\Project\PaidDevelopment\07_PDF_Author-Email-Extractor_begin_￥500\data'
    excel = r'C:\Users\fanyu\Desktop\Project\PaidDevelopment\07_PDF_Author-Email-Extractor_begin_￥500\output.xls'

    # add sheets
    print('add sheets...')
    try:
        handle_folder(base_dir, excel)
    except Exception as e:
        raise e
    print('run parser...')
    files = [f for f in os.listdir(base_dir) if
             os.path.isfile(os.path.join(base_dir, f)) and os.path.splitext(f)[1] == '.txt']
    for index, file in enumerate(files):
        full_path = os.path.join(base_dir, file)
        print('parsing file:\t', index + 1, '/', len(files), '\t', file)
        author_info = AuthorInfo2(full_path)
        try:
            author = author_info.run()
        except Exception as e:
            raise e
        if isinstance(author, list):
            if len(file) > 20:
                sheet_name = file[:20] + '...'
            else:
                sheet_name = file
            try:
                sheet_name = '{:0>4d}_{}'.format(index, sheet_name)
                save2Excel(excel, sheet_name, author)
            except Exception as e:
                raise e

# Script for convert Firefox history from .sqlite to .xls(x) file.
# author Ilya Matthew Kuvarzin <luceo2011@yandex.ru>
# version 1.0 dated January 14, 2020

import sqlite3
import datetime
import re
import openpyxl
from openpyxl.styles import Font, Border, Side
import sys
import os
import math

# settings
file_name = 'History'
rows_limit = 25000
file_extension = '.xlsx'


def cls():
    os.system('cls' if os.name == 'nt' else 'clear')


def initialization():
    book = openpyxl.Workbook()
    sh = book.active
    sh.title = 'History'
    sh['A1'] = 'Date'
    sh['B1'] = 'Url'
    sh['C1'] = 'Title'
    sh['D1'] = 'Description'
    return book


def save(book, ind, n):
    if n:
        num_text = '_' + str(n)
    else:
        num_text = ''
    cls()
    font = Font(bold=True, size=14)
    border = Side(style='thin', color='000000')
    for i in range(1, ind):
        for j in range(1, 5):
            cel = sheet.cell(i, j)
            if i == 1:
                cel.font = font
            if (i != 1) and (j == 2):
                cel.style = 'Hyperlink'
            cel.border = Border(left=border, top=border, right=border, bottom=border)
        pro = i/ind
        sys.stdout.write(
            '\rDecoration: [' + '#' * math.floor(pro * 25) + '_' * math.floor((1 - pro) * 25) + '] '
            + str(round(pro * 10000) / 100) + '%')
        sys.stdout.flush()
    sys.stdout.write('\rSaving file ' + file_name + num_text + '...')
    sys.stdout.flush()
    book.save(file_name + num_text + file_extension)
        
        
conn = sqlite3.connect("places.sqlite")
cursor = conn.cursor()
wb = initialization()
sheet = wb.active
index = 2
result = cursor.execute('''SELECT H.visit_date, P.url, P.title, P.description
                        FROM moz_places as P
                                LEFT OUTER JOIN moz_historyvisits as H
                                ON P.id = H.place_id
                        ORDER BY H.visit_date DESC''').fetchall()
rowcount = len(result)
if rowcount > rows_limit:
    number = 1
else:
    number = False
for row in result:
        if row[0] is not None:
            sheet['A' + str(index)] = datetime.datetime.fromtimestamp(float(row[0])/1000000)
        else:
            sheet['A' + str(index)] = '???'
        match = re.search('https?://w*\.?', row[1])
        if match:
            site = row[1].replace(match.group(), '')
            site = site.replace(re.search('/.*', site).group(), '')
            sheet['B' + str(index)].hyperlink = row[1]
            sheet['B' + str(index)].value = site
        else:
            sheet['B' + str(index)] = row[1]
        sheet['C' + str(index)] = row[2]
        sheet['D' + str(index)] = row[3]
        if (rowcount - (number-1)*rows_limit) > rows_limit:
            count = rows_limit
        else:
            count = rowcount - (number-1)*rows_limit
        if index == rows_limit:
            save(wb, index, number)
            number += 1
            index = 2
            wb = initialization()
            sheet = wb.active
        else:
            progress = index/count
            sys.stdout.write('\rProgress: [' + '#' * math.floor(progress*25) + '_' * math.floor((1-progress)*25) + '] '
                             + str(round(progress*10000)/100) + '%')
            sys.stdout.flush()
            index += 1

if index > 2:
    save(wb, index, number)

conn.close()



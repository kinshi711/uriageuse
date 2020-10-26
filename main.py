import xlwings as xw
import os
import numpy as np

def osibori_sum():
    o_sum = sum(
        [w_sheet.range('B{}'.format(n1)).options(empty=0).value * w_sheet.range('C{}'.format(n1)).options(empty=0).value
    for n1 in range(1, 5, 1)])
    return o_sum
def water_sum():
    w_sum = sum(
        [w_sheet.range('B{}'.format(n2)).options(empty=0).value * w_sheet.range('C{}'.format(n2)).options(empty=0).value
    for n2 in range(5, 10, 1)])
    return w_sum
def ice_sum():
    i_sum = sum(
        [w_sheet.range('B{}'.format(n3)).options(empty=0).value * w_sheet.range('C{}'.format(n3)).options(empty=0).value
    for n3 in range(10, 17, 1)])
    return i_sum
def day_sum():
    osibori = osibori_sum()
    water = water_sum()
    ice = ice_sum()
    d_sum = osibori + water + ice
    return d_sum
def month_sum():
    for m in range(1, 13, 1):
        m_sheet = s_book.sheets['{}月'.format(month)]
        m_list = m_sheet.range('E2:E32').options(empty=0).value
        m_sum = sum(m_list)
    return m_sum
def year_sum():
    y_sheet = s_book.sheets['年'].range('B2:B13').options(empty=0).value
    y_sum = sum(y_sheet)
    return y_sum
path_dict = {}
dir = os.getcwd()
path_dict['w_path'] = os.path.join(dir, '書き込み用エクセルファイル.xlsm')
w_book = xw.Book(path_dict['w_path'])

month = int(w_book.sheets['売り上げ記入用'].range('H2').options(empty='選択肢してください').value)
day = int(w_book.sheets["売り上げ記入用"].range('J2').options(empty='選択して下さい').value)
s_dir = os.path.join(dir, '保存先')
year = int(w_book.sheets['売り上げ記入用'].range('F2').options(empty='選択肢してください').value)
path_dict['s_path'] = os.path.join(s_dir, '{}年保存先.xlsx'.format(year))
path_dict['original_path'] = os.path.join(s_dir, 'original.xlsx')
save_path = os.path.join(s_dir, '{}年保存先.xlsx'.format(year))
original_book = xw.Book(path_dict['original_path'])

sheet_list = ('年', '12月', '11月', '10月', '9月', '8月', '7月', '6月', '5月', '4月', '3月', '2月', '1月')
original_sheet = original_book.sheets
original_year_sheet = original_book.sheets['年']
original_year_value = np.array(original_year_sheet.range('A1:C13').options(empty=0).value)

try:
    s_path_book= xw.Book(path_dict['s_path'])
    s_path_book.close()
except FileNotFoundError:
    s_book = xw.Book()
    s_book.save(path_dict['s_path'])
    for sheet_tapple in sheet_list:
        s_book.sheets.add(name='{}'.format(sheet_tapple))
        s_book.sheets['年'].range('A1:F13').options(empty=0).value= original_year_value
    for month_ in (1,2,3,4,5,6,7,8,9,10,11,12):
        s_book.sheets['{}月'.format(month_)].range('A1:F32').options(empty=0).value = original_sheet['{}月'.format(month_)].range('A1:F32').options(empty=0).value
        s_book.save(path_dict['s_path'])

w_sheet= w_book.sheets['売り上げ記入用']
osibori = osibori_sum()
water = water_sum()
ice = ice_sum()
sum_day = day_sum()
day_1 = 0
for i in range(1,32):
    if str(i) in str(day):
        day_1 = i + 1
s_book = xw.Book(path_dict['s_path'])
s_book.sheets['{}月'.format(month)].range('B{}'.format(day_1)).options(empty=0).value = osibori
s_book.sheets['{}月'.format(month)].range('C{}'.format(day_1)).options(empty=0).value = water
s_book.sheets['{}月'.format(month)].range('D{}'.format(day_1)).options(empty=0).value = ice
s_book.sheets['{}月'.format(month)].range('E{}'.format(day_1)).options(empty=0).value = sum_day

sum_month = month_sum()
s_book.sheets['{}月'.format(month)].range('F{}'.format(day_1)).value = sum_month

for i in range(1,13):
    if str(i) in str(month):
        month_1 = i + 1

s_book.sheets['年'].range('B{}'.format(month_1)).value = sum_month

sum_year = year_sum()
s_book.sheets['年'].range('C{}'.format(month_1)).value = sum_year
w_book.close()
original_book.close()
s_book.save()
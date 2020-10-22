import xlwings as xw
import os
import numpy as np


path_dict = {}  # 空辞書を作って　パスを追加していく
dir = os.getcwd()  # 絶対参照というか自分の今いる場所がわかる /uriage
path_dict['w_path'] = os.path.join(dir, '書き込み用エクセルファイル.xlsm')  # os.path.joinっているのはパスをつなげるときに必要、joinだけだとできない
w_book= xw.Book(path_dict['w_path'])  # 書き込み用のエクセルファイルを開く


month = int(w_book.sheets['売り上げ記入用'].range('G1').options(empty=0).value)  # 書き込まれた月の参照
day = int(w_book.sheets["売り上げ記入用"].range('I1').options(empty=0).value)  # 書き込まれた日の参照
s_dir = os.path.join(dir, '保存先')  # 保存用フォルダーの参照
year = int(w_book.sheets['売り上げ記入用'].range('F2').options(empty=0).value)  # 書き込まれた年度を参照
path_dict['s_path'] = os.path.join(s_dir, '{}年保存先.xlsx'.format(year))  # 保存用エクエルファイルのパス
path_dict['original_path'] = os.path.join(s_dir, 'shuya_natural保存先.xlsx')  # 原本のエクセルファイルのパス
path_dict['2020.xlsx'] = os.path.join(s_dir, '2020年保存先.xlsx')# 2020年のエクセルファイル作成時のpathを呼び出すのを追加
save_path = os.path.join(s_dir, '{}年保存先.xlsx'.format(year))
wb_2020 = xw.Book()
original_book = xw.Book(path_dict['original_path'])



# if path_dict['2020.xlsx'] in s_dir:
#     copy_value=np.array(original_book.sheets[num].range('A1:F32').options(empty=0).value for num in (0,14,1))
#     wb_2020.close()
# else:

sheet_list = ('年','12月','11月','10月','9月','8月','7月','6月','5月','4月','3月','2月','1月')
original_sheet= original_book.sheets
original_year_sheet = original_book.sheets['年']
original_year_value =np.array(original_year_sheet.range('A1:C13').options(empty=0).value)
try:
    book2020 =xw.Book(path_dict['2020.xlsx'])
    book2020.close()
except FileNotFoundError:
    wb_2020.save(path_dict['2020.xlsx'])
    for sheet_tapple in sheet_list:
        wb_2020.sheets.add(name='{}'.format(sheet_tapple))
        wb_2020.sheets['年'].range('A1:F13').options(empty=0).value = original_year_value
    for month_ in (1,2,3,4,5,6,7,8,9,10,11,12):
        wb_2020.sheets['{}月'.format(month_)].range('A1:F32').options(empty=0).value = original_sheet['{}月'.format(month_)].range('A1:F32').options(empty=0).value
    wb_2020.save(path_dict['2020.xlsx'])

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

w_sheet= w_book.sheets['売り上げ記入用']# 売り上げ記入用
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
osibori = osibori_sum()
water = water_sum()
ice = ice_sum()
sum_day = day_sum()
day_1 = 0

for i in range(1,32):
    if str(i) in str(day):
        day_1 = i + 1

wb_2020= xw.Book(path_dict['2020.xlsx'])
wb_2020.sheets['{}月'.format(month)].range('B{}'.format(day_1)).value = osibori
wb_2020.sheets['{}月'.format(month)].range('C{}'.format(day_1)).value = water
wb_2020.sheets['{}月'.format(month)].range('D{}'.format(day_1)).value = ice
wb_2020.sheets['{}月'.format(month)].range('E{}'.format(day_1)).value = sum_day

#wb_2020.sheets['{}月'.format(month_)].range('A1:F32').options(empty=0).value = original_sheet['{}月'.format(month_)].range('A1:F32').options(empty=0).value
def month_sum():
    for m in range(1, 13, 1):
        m_sheet = wb_2020.sheets['{}月'.format(month)]
        m_list = m_sheet.range('E2:E32').options(empty=0).value
        m_sum = sum(m_list)
    return m_sum
sum_month = month_sum()
wb_2020.sheets['{}月'.format(month)].range('F{}'.format(day_1)).value = sum_month

for i in range(1,13):
    if str(i) in str(month):
        month_1 = i + 1

wb_2020.sheets['年'].range('B{}'.format(month_1)).value = sum_month

def year_sum():
    y_sheet = wb_2020.sheets['年'].range('B2:B13').options(empty=0).value
    y_sum = sum(y_sheet)
    return y_sum
sum_year = year_sum()
wb_2020.sheets['年'].range('C{}'.format(month_1)).value = sum_year
w_book.close()
original_book.close()
wb_2020.save()
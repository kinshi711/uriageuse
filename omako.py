

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

wb_2020 = xw.Book()
original_book = xw.Book(path_dict['original_path'])

sheet_list = ('年','12月','11月','10月','9月','8月','7月','6月','5月','4月','3月','2月','1月')
for sheet_tapple in sheet_list:
    wb_2020.sheets.add(name='{}'.format(sheet_tapple))
wb_2020.save(path_dict['2020.xlsx'])
original_sheet= original_book.sheets
original_year_sheet = original_book.sheets['年']
original_year_value =np.array(original_year_sheet.range('A1:C13').options(empty=0).value)
wb_2020.sheets['年'].range('A1:F13').options(empty=0).value = original_year_value
for month_ in (1,2,3,4,5,6,7,8,9,10,11,12):
    wb_2020.sheets['{}月'.format(month_)].range('A1:F32').options(empty=0).value = original_sheet['{}月'.format(month_)].range('A1:F32').options(empty=0).value
    wb_2020.save(path_dict['2020.xlsx'])
wb_2020.save(path_dict['2020.xlsx'])

if   path_dict['s_path'] in s_dir:  # 新規保存用のエクセルファイルがフォルダーに既に有ったら　
     s_book= xw.Book()  # 新規保存用のエクセルファイル
     s_book.close()# 新規でエクセルファイルもいらないので消す
elif path_dict['w_path'] in path_dict['s_path']:
    s_book= xw.Book()  # 新規保存用のエクセルファイル
    s_book.close()
else:
    s_book= xw.Book()  # 新規保存用のエクセルファイル
    try:
        s_book.save(path_dict['{}.xlsx'.format(year)])
        s_book.close()#保存用のファイルがフォルダーに追加されたので消す。
    except:
        s_book.close()

w_sheet= w_book.sheets['売り上げ記入用']# 売り上げ記入用


def osibori_sum():
    o_sum = sum(
        [w_sheet.range('B{}'.format(n1)).options(empty=0).value * w_sheet.range('C{}'.format(n1)).options(empty=0).value
         for n1 in range(1, 5, 1)])
def water_sum():
    w_sum = sum(
        [w_sheet.range('B{}'.format(n2)).options(empty=0).value * w_sheet.range('C{}'.format(n2)).options(empty=0).value
         for n2 in range(5, 10, 1)])

def ice_sum():
    i_sum = sum(
        [w_sheet.range('B{}'.format(n3)).options(empty=0).value * w_sheet.range('C{}'.format(n3)).options(empty=0).value
         for n3 in range(10, 17, 1)])

def day_sum():
    osibori = osibori_sum()
    water = water_sum()
    ice = ice_sum()
    d_sum = osibori + water + ice


def month_sum():
    for m in range(1, 13, 1):
        m_sheet = s_book.sheets.range('{}月'.format(m)).value
        m_sum = sum([m_sheet.range('E{}'.format(num)).options(empty=0).value for num in (2, 32, 1)])
        return m_sum


def year_sum():
    y_sheet = s_book.sheets['年'].range('B2:B13').options(empty=0).value[0:][0:]
    y_sum = sum(y_sheet)
    return y_sum

w_book.close()
original_book.close()
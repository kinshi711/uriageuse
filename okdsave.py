import os
import xlwings as xw
import numpy as np

dirname = os.getcwd()  # 現在の場所を調べる
save_folder_path = os.path.join(dirname, '保存先')
w_path = os.path.join(dirname, '書き込み用エクセルファイル.xlsm')  # 書き込みエクセルファイルのpath
natural_save_excel = os.path.join(save_folder_path, 'shuya_natural保存先.xlsx')  # 原本のエクセルファイルのpath
path_dict = {}  # 空辞書を作って　パスを作っておく

w_Book = xw.Book(w_path)  # 書き込みエクセルファイルの参照
w_year = int(w_Book.sheets['売り上げ記入用'].range('F2').options(empty=2020).value)  # 書き込みエクセルファイルに書かれた年度を参照
path_dict['2020.xlsx'] = os.path.join(save_folder_path, '2020年保存先.xlsx')
# path_dictは　2020年のエクセルファイル作成時のpathを呼び出す。

if path_dict['2020.xlsx'] in save_folder_path:
    wb2020 = xw.Book(path_dict['2020.xlsx'])
    for m in range(1, 13, 1):
        natural_month_sheet = xw.Book(natural_save_excel).sheets['{}月'.format(m)]
        natural_year_sheet = xw.Book(natural_save_excel).sheets('年')
        copy_month_value = natural_month_sheet.range('A1:F32').options(empty=0).value
        copy_year_value = natural_year_sheet.range('A1:C13')

else:
    wb2020 = xw.Book()  # 2020年保存先のエクセルファイルの作成
    wb2020.save(path_dict['2020.xlsx'])  # 保存先の指
    for m in range(1, 13, 1):
        natural_month_sheet = xw.Book(natural_save_excel).sheets['{}月'.format(m)]
        natural_year_sheet = xw.Book(natural_save_excel).sheets('年')
        copy_month_value = natural_month_sheet.range('A1:F32').options(empty=0).value
        copy_year_value = natural_year_sheet.range('A1:C13').options(empty=0).value
        wb2020.sheets.add(name='{}月'.format(m))
    wb2020.save()
    wb2020.close()

path_dict['{}.xlsx'.format(w_year)] = os.path.join(save_folder_path, '{}年保存先.xlsx'.format(w_year))
year_Book = xw.Book()  # エクセルを作成
if path_dict['{}.xlsx'.format(w_year)] in save_folder_path:
    year_Book.close()
else:
    # 書き込みエクセルファイルの年を取得してそれをもとにエクセルファイル名を取得してその上でそこのパスを作成している
    year_Book.save(path_dict['{}.xlsx'.format(w_year)])  # pathを指定して保存する
    year_Book.close()  # 書かれた年の作成ファイルを閉じる


w_Book.close()  # 書き込みファイル閉じる
xw.Book(natural_save_excel).close()  # 原本ファイル閉じる
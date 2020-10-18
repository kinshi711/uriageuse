import os
import xlwings as xw

dirname = os.getcwd() #現在の場所を調べる
save_folder_path = os.path.join(dirname , '保存先')
w_path = os.path.join(dirname,'書き込み用エクセルファイル.xlsm') #書き込みエクセルファイルのpath
natural_save_excel = os.path.join(save_folder_path,'shuya_natural保存先.xlsx') #原本のエクセルファイルのpath
path_dict = {}#空辞書を作って　パスを作っておく

w_Book = xw.Book(w_path)# 書き込みエクセルファイルの参照
w_year = int(w_Book.sheets['売り上げ記入用'].range('F2').options(empty=2020).value)#書き込みエクセルファイルに書かれた年度を参照
w_month = str(w_Book.sheets['売り上げ記入用'].range('G1').value)
w_Book.close()#閉じまちた

path_dict['{}.xlsx'.format(w_year)] = os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))
#path_dictは　2020年のエクセルファイル作成時のpathを呼び出す。
s_book = xw.Book(os.path.join(save_folder_path, '{}年保存先.xlsx'.format(w_year)))
original = xw.Book(natural_save_excel).sheets.active
copy_values = original.range('A1:F32').value#リスト上で値はとれている
s_book = s_book.sheets['{}月'.format(w_month)].range('A1:F32').value
s_book = copy_values.copy(destination=s_book)
s_book.save


def make_s_book():
    s_book = xw.Book(os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year)))
    original = xw.Book(natural_save_excel).sheets.active
    copy_values = original.range('A1:F32').value
    s_book = s_book.sheets['{}月'.format(w_month)].range('A1:F32')
    save_book = copy_values.copy(destination = s_book)

    print(copy_values)

make_s_book()
# make_s_book()
# for r in range(1, 32):
#     for c in range(1, 7):
#         copy_value = original.cells(row=r, column=c).value
#
#         s_book.cells(row=r, column=c, value=copy_value)
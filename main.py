import os
import xlwings as xw


dirname= os.getcwd() #絶対参照というか自分の今いる場所がわかる /uriage
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')#os.path.joinっているのはパスをつなげるときに必要、joinだけだとできない

w_Book = xw.Book(w_path)#書き込み用エクセルファイルのパスのファイルを読み込んで、それをw_bookと置いて活用しやすいようにする

#書き込み用エクセルファイルは開けている
w_year =str(w_Book.sheets['売り上げ記入用'].range('F2').value)#書き込み用エクセルファイルで選択された年の値を読み取る
w_month = str(w_Book.sheets['売り上げ記入用'].range('G1').value)
w_day = str(w_Book.sheets["売り上げ記入用"].range('I1').value)

save_folder_path = os.path.join(dirname,'保存先')
s_path = os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))

xw1 = xw.Book(w_path)

w_year =int(xw1.sheets['売り上げ記入用'].range('F2').options(empty=2020).value)
s_path= os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))
xw2 = xw.Book(s_path)
sheet1 = xw1.sheets['売り上げ記入用']

def osi_sum():
    o_sum = sum([sheet1.range('B{}'.format(n1)).options(empty=0).value * sheet1.range('C{}'.format(n1)).options(empty=0).value for n1 in range(1, 5, 1)])
    return o_sum
def water_sum():
    w_sum = sum([sheet1.range('B{}'.format(n2)).options(empty=0).value * sheet1.range('C{}'.format(n2)).options(empty=0).value for n2 in range(5, 10, 1)])
    return w_sum
def ice_sum():
    i_sum = sum([sheet1.range('B{}'.format(n3)).options(empty=0).value * sheet1.range('C{}'.format(n3)).options(empty=0).value for n3 in range(10, 17, 1)])
    return i_sum
def day_sum():
    osibori = osi_sum()
    water = water_sum()
    ice = ice_sum()
    d_sum = osibori + water + ice
    return d_sum
osibori = osi_sum()
water = water_sum()
ice = ice_sum()
day = day_sum()

def main_to_save():
    xw2.sheets.active.range('B2').value(osibori)
main_to_save()

def month_sum():
    for m in range(1,13,1):
        sheet2 = xw2.sheets.range('{}月'.format(m)).options(empty=0).value
        m_sum= sum([sheet2['E{}'.format(num)].value for num in (2,32,1)])
        return m_sum
def year_sum():
    y_sheet = xw2.sheets['年'].range('B2:B13').options(empty=0).value[0:][0:]
    y_sum = sum(y_sheet)
    return y_sum


xw1.close()
xw2.close()

natural_save_excel = os.path.join(save_folder_path, 'shuya_natural保存先.xlsx')
path_dict = {}  # 空辞書を作って　パスを作っておく

w_Book = xw.Book(w_path)  # 書き込みエクセルファイルの参照
w_year = int(w_Book.sheets['売り上げ記入用'].range('F2').options(empty=2020).value)  # 書き込みエクセルファイルに書かれた年度を参照
path_dict['2020.xlsx'] = os.path.join(save_folder_path, '2020年保存先.xlsx')
# path_dictは　2020年のエクセルファイル作成時のpathを呼び出す。

if path_dict['2020.xlsx'] in save_folder_path:
    wb2020 = xw.Book(path_dict['2020.xlsx'])
    wb2020.close()
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



#    s_Book.sheets['{}月'.format(w_month)].range('B{}'.format(w_month))


import os
from openpyxl import Workbook
import xlwings as xw
import shuya_sum
import shuya_save

day_sum = syuya.day_sum()#日合計の関数を取得
month_sum = syuya.month_sum()#月合計の関数を取得
year_sum = syuya.year_sum()#年合計の関数を取得
wet_towel = syuya.osibori#おしぼりのインスタンス化されている関数を取得
water = syuya.water#水のインスタンス化されている関数を取得
ice = syuya.ice#氷の
# インスタンス化されている関数を取得
na = syuya.na
dirname= os.getcwd() #絶対参照というか自分の今いる場所がわかる /uriage
#print(dirname) /uriageになる
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')#os.path.joinっているのはパスをつなげるときに必要、joinだけだとできない

w_Book = xw.Book(w_path)#書き込み用エクセルファイルのパスのファイルを読み込んで、それをw_bookと置いて活用しやすいようにする
w_year =str(w_Book.sheets['売り上げ記入用'].range('F2').value)#書き込み用エクセルファイルで選択された年の値を読み取る
w_month = str(w_Book.sheets['売り上げ記入用'].range('G1').value)
w_day = str(w_Book.sheets["売り上げ記入用"].range('I1').value)
def sum_to_save():
    wb = xw.book()
    save_folder_path = os.path.join(dirname, '保存先')  # ここで売り上げの中の保存先っているディレクトリに行ってねって指示できている
    save_file = os.path.join(save_folder_path,'{}念保存先.xlsx'.format(w_year))
    save_month_sheet = save_file.sheets['{}月'.format(w_month)]
    save_day = save_month_sheet.Range('B{}'.format(w_day + 1))
    save_day.value = day_sum



#今の段階の整理
#まず、合計ファイルの合計の関数の結果は変数に置くことができた。あとはこれを保存用に渡せばいい
#どうやったら渡せるのか、まずはしゅうやのファイルでどこの部分が保存用を表しているのかを探す
#day_sumの日合計の値を今から保存用エクセルファイルに移す
#path_dict = {'{}.xlsx'.format(w_year):os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))}#書き込み用エクセルファイルかかれた


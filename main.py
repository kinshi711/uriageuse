import os
from openpyxl import Workbook
import xlwings as xw
from shuya_sum import osi_sum,water_sum,ice_sum,month_sum,year_sum
import shuya_save


dirname= os.getcwd() #絶対参照というか自分の今いる場所がわかる /uriage
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')#os.path.joinっているのはパスをつなげるときに必要、joinだけだとできない

w_Book = xw.Book(w_path)#書き込み用エクセルファイルのパスのファイルを読み込んで、それをw_bookと置いて活用しやすいようにする

#書き込み用エクセルファイルは開けている
w_year =str(w_Book.sheets['売り上げ記入用'].range('F2').value)#書き込み用エクセルファイルで選択された年の値を読み取る
w_month = str(w_Book.sheets['売り上げ記入用'].range('G1').value)
w_day = str(w_Book.sheets["売り上げ記入用"].range('I1').value)

save_folder_path = os.path.join(dirname,'保存先')
s_path = os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))
s_Book = xw.Book(s_path)
def main_to_save():
    a = osi_sum()
    print(a)

main_to_save()
#    s_Book.sheets['{}月'.format(w_month)].range('B{}'.format(w_month))





#今の段階の整理
#まず、合計ファイルの合計の関数の結果は変数に置くことができた。あとはこれを保存用に渡せばいい    def sum_to_save():
 #       wb = xw.Book()
  #      save_folder_path = os.path.join(dirname, '保存先')  # ここで売り上げの中の保存先っているディレクトリに行ってねって指示できている
   #     save_file = os.path.join(save_folder_path, '{}念保存先.xlsx'.format(w_year))
    #    save_month_sheet = save_file.sheets['{}月'.format(w_month)]  # sheet の定義がされてないよ
     #   save_day = save_month_sheet.Range('B{}'.format(w_day + 1))  # I1+1になってしまってるよ

#どうやったら渡せるのか、まずはしゅうやのファイルでどこの部分が保存用を表しているのかを探す
#day_sumの日合計の値を今から保存用エクセルファイルに移す
#path_dict = {'{}.xlsx'.format(w_year):os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))}#書き込み用エクセルファイルかかれた


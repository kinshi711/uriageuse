import os
import xlwings as xw

dirname= os.getcwd() #現在の場所を調べる
save_folder_path =os.path.join(dirname,'保存先')
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm') #書き込みエクセルファイルのpath
natural_save_excel = os.path.join(dirname,'shuya_natural保存先.xlsx') #原本のエクセルファイルのpath
path_dict = {}#空辞書を作って　パスを作っておく

w_Book = xw.Book(w_path)# 書き込みエクセルファイルの参照
w_year = int(w_Book.sheets['売り上げ記入用'].range('F2').options(empty=2020).value)#書き込みエクセルファイルに書かれた年度を参照
w_Book.close()#閉じまちた

path_dict['2020.xlsx'] = os.path.join(save_folder_path,'2020年保存先.xlsx')
#path_dictは　2020年のエクセルファイル作成時のpathを呼び出す。

wb_2020 = xw.Book()# 2020年保存先のエクセルファイルの作成
wb_2020.save(path_dict['2020.xlsx'])#保存先の指定
wb_2020.close()#次回スタートする時にバグらないように停止

path_dict['{}.xlsx'.format(w_year)] = os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))
#書き込みエクセルファイルの年を取得してそれをもとにエクセルファイル名を取得してその上でそこのパスを作成している
year_Book = xw.Book()#エクセルを作成
new_year_Book =year_Book.save(path_dict['{}.xlsx'.format(w_year)])#pathを指定して保存する　
year_Book.close()#エクセルを閉じている
# sht.range('A1').options(empty='NA').value
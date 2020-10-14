import os
import datetime
import xlwings as xw
dirname= os.getcwd() #絶対参照というか自分の今いる場所がわかる /uriage
#print(dirname) /uriageになる
save_folder_path =os.path.join(dirname,'保存先')#ここで売り上げの中の保存先っているディレクトリに行ってねって指示できている
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')
natural_save_excel = os.path.join(dirname,'shuya_natural保存先.xlsx')
w_Book = xw.Book(w_path)
w_year =int(w_Book.sheets['売り上げ記入用'].range('F2').value)
w_Book.close()
path_dict={'save_folder' : os.path.join(save_folder_path),'2020.xlsx':os.path.join(save_folder_path,'2020年保存先.xlsx')}
wb_2020 = xw.Book()
wb_2020.save(path_dict['2020.xlsx'])
wb_2020.close()
t_2020 = os.path.getctime(path_dict['2020.xlsx'])
d_2020 = datetime.datetime.fromtimestamp(t_2020)
year_2020 = d_2020.year
save_file_path_2020 = os.path.join(save_folder_path,'{}年保存先'.format(year_2020))
save_file_path_2020_excel = os.path.join(save_file_path_2020,'.xlsx')
path_dict = {'{}.xlsx'.format(w_year):os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))}
new_Book = xw.Book()
year_Book =new_Book.save(path_dict['{}.xlsx'.format(w_year)])
new_Book.close()


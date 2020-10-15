import os
import datetime
import xlwings as xw
dirname= os.getcwd() #絶対参照というか自分の今いる場所がわかる /uriage
#print(dirname) /uriageになる
save_folder_path =os.path.join(dirname,'保存先')#ここで売り上げの中の保存先っているディレクトリに行ってねって指示できている
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')#os.path.joinっているのはパスをつなげるときに必要、joinだけだとできない
natural_save_excel = os.path.join(save_folder_path,'shuya_natural保存先.xlsx')#しゅうやの保存用原本にアクセスしている
w_Book = xw.Book(w_path)#書き込み用エクセルファイルのパスのファイルを読み込んで、それをw_bookと置いて活用しやすいようにする
w_year =str(w_Book.sheets['売り上げ記入用'].range('F2').value)#書き込み用エクセルファイルで選択された年の値を読み取る
w_Book.close()#いったん閉じる　無駄なリソースを使わないために
path_dict={'save_folder' : os.path.join(save_folder_path),'2020.xlsx':os.path.join(save_folder_path,'2020年保存先.xlsx')}
#save_folderって入力したら後のやつが呼び出されるsave_folder_pathがでて、2020.xlsxだと2020年保存先.xlsxが出てくる　保存先ディレクトリの中にある
wb_2020 = xw.Book()#今稼働しているやつをwb_2020としている
wb_2020.save(path_dict['2020.xlsx'])#さっきの2020年保存先.xlsxの中身をwb_2020に保存する
wb_2020.close()
t_2020 = os.path.getctime(path_dict['2020.xlsx'])#getctimeとは　そのファイルが最後に更新された時の日時を取得するメソッド
d_2020 = datetime.datetime.fromtimestamp(t_2020)#fromtimestampはスタンプ値を日本の表記にしている
year_2020 = d_2020.year #先ほど直した表記から年の値を読み取る

save_file_path_2020 = os.path.join(save_folder_path,'{}年保存先'.format(year_2020))#保存したいエクセルファイルのあるディレクトリを指定してそのなかにあるファイルを参照するための年数を指定している　ここでは迄エクセルファイルにアクセルしていない
save_file_path_2020_excel = os.path.join(save_file_path_2020,'.xlsx')#先ほどのディレクトリから読み込みたいエクセルファイルを指定する
path_dict = {'{}.xlsx'.format(w_year):os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))}#書き込み用エクセルファイルかかれた
#年と対応する年の保存用エクセルファイルを読み込む
new_Book = xw.Book()#新しいエクセルファイルを作る
year_Book =new_Book.save(path_dict['{}.xlsx'.format(w_year)])#書き込み用エクセルファイルの記入された年に対応する保存用のエクセルファイルをつくる
#これは２０２１年とかのよう
new_Book.close()

w_Book = xw.Book(w_path)#アクティブアプリケーションのブック管理インスタンスを返す,つまり今動いているエクセルをw_bookと命名
w_year =w_Book.sheets['売り上げ記入用'].range('F2').value#書き込み用エクセルファイルに記入された年をw_yearにしている
year_dict = {'{}年.xlsx'.format(w_year): os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))}
#w_year年.xlsxっていうのがきたら同じ年の保存先ファイルを読み込む
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')
#書き込み用エクセルファイルのパスを指定
s_path= os.path.join(year_dict['{}年.xlsx'.format(w_year)])#保存先エクセルファイルのパスを変数に置く
xw1 = xw.Book(w_path)#書き込み用エクセルファイルを変数に置く
xw2 = xw.Book(s_path)#保存先エクセルファイルを変数に置く
sheet1 = xw1.sheets['売り上げ記入用']#書き込み用エクセルファイルの売り上げ記入用シートをsheet1遠く

try:#例外処理の対応をするために使う
    def day_sum():
        def osi_sum():#おしぼりの合計の値を調べる　なんで値をforで回していないかっていうと、エクセルの使用でできないらしい。　だから仮の文字を入れておいて、それを文字で回す
            o_sum = sum([sheet1['B{}'.format(n1)].value * sheet1['C{}'.format(n1)].value for n1 in range(1, 5, 1)])#1-4まで
            return  o_sum
        def water_sum():#水の合計の値を示している
            w_sum = sum([sheet1['B{}'.format(n2)].value * sheet1['C{}'.format(n2)].value for n2 in range(5, 10, 1)])
            return w_sum
        def ice_sum():#氷の合計の値を調べている
            i_sum = sum([sheet1['B{}'.format(n3)].value * sheet1['C{}'.format(n3)].value for n3 in range(10, 17, 1)])
            return i_sum
        osibori = osi_sum()#インスタンス化
        water= water_sum()
        ice = ice_sum()
        d_sum =osi_sum()+water_sum()+ice_sum()#一日の合計を示すものをインスタンス化している
    def month_sum(command):
        for m in range(1,13,1):
            sheet2 = xw2.sheets['{}月'.format(m)]#sheet2に保存用エクエルファイルの各月を対応するようにしている
            m_sum= sum(sheet2['E{}'.format(num)].value for num in range(2,30,1))
    def year_sum(self,command):#
            y_sheet = xw2.sheets['年'].range('B2:B13').value[0:][0:]
            y_sum = sum(self.y_sheet)
except TypeError:
    print('数字が書かれてない場合があります　何もないところは0を書いてください')
xw1.close()
xw2.close()

def na():
    print("ari")
na = na()
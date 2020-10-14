import os
import datetime
import xlwings as xw
w_Book = xw.Book(w_path)
w_year =w_Book.sheets['売り上げ記入用'].range('F2').value
year_dict = {'{}年.xlsx'.format(w_year): os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))}
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')
s_path= os.path.join(year_dict['{}年.xlsx'.format(w_year)])
xw1 = xw.Book(w_path)
xw2 = xw.Book(s_path)
sheet1 = xw1.sheets['売り上げ記入用']
try:
    def day_sum():
        def osi_sum():
            o_sum = sum([sheet1['B{}'.format(n1)].value * sheet1['C{}'.format(n1)].value for n1 in range(1, 5, 1)])
        def water_sum():
            w_sum = sum([sheet1['B{}'.format(n2)].value * sheet1['C{}'.format(n2)].value for n2 in range(5, 10, 1)])
        def ice_sum():
            i_sum = sum([sheet1['B{}'.format(n3)].value * sheet1['C{}'.format(n3)].value for n3 in range(10, 17, 1)])
        osibori = osi_sum()
        water= water_sum()
        ice = ice_sum()
        d_sum =osi_sum()+water_sum()+ice_sum()
    def month_sum(command):
        for m in range(1,13,1):
            sheet2 = xw2.sheets['{}月'.format(m)]
            m_sum= sum([sheet2['E{}'.format(num)].value+sheet2['E{}'.format(num)].value for num in range(2,30,1)])
    def year_sum(self,command):
            y_sheet = xw2.sheets['年'].range('B2:B13').value[0:][0:]
            y_sum = sum(self.y_sheet)
except TypeError:
    print('数字が書かれてない場合があります　何もないところは0を書いてください')
xw1.close()
xw2.close()
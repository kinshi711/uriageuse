import xlwings as xw
import os

dirname= os.getcwd() #現在の場所を調べる
save_folder_path =os.path.join(dirname,'保存先')
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')
w_year =int(w_Book.sheets['売り上げ記入用'].range('F2').options(empty=2020).value)
s_path= os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))
xw1 = xw.Book(w_path)
xw2 = xw.Book(s_path)
sheet1 = xw1.sheets['売り上げ記入用']

def day_sum():
    def osi_sum():
        o_sum = sum([sheet1.range('B{}'.format(n1)).options(empty=0).value * sheet1.range('C{}'.format(n1)).options(empty=0).value for n1 in range(1, 5, 1)]           return o_sum
    def water_sum():
        w_sum = sum([sheet1.range('B{}'.format(n2)).options(empty=0).value * sheet1.range('C{}'.format(n2)).options(empty=0).value for n2 in range(5, 10, 1)])
        return w_sum
    def ice_sum():
        i_sum = sum([sheet1.range('B{}'.format(n3)).options(empty=0).value * sheet1.range('C{}'.format(n3)).options(empty=0).value for n3 in range(10, 17, 1)])
        return i_sum
    osibori = osi_sum()
    water = water_sum()
    ice = ice_sum()
    d_sum = osibori+water+ice
    return d_sum
def month_sum():
    for m in range(1,13,1):
        sheet2 = xw2.sheets.range('{}月'.format(m)).options(empty=0).value
        m_sum= sum([sheet2['E{}'.format(num)].value for num in (2,32,1)])
        return m_sum
def year_sum():
    y_sheet = xw2.sheets['年'].range('B2:B13').options(empty=0).value[0:][0:]
    y_sum = sum(y_sheet)
    return y_sum



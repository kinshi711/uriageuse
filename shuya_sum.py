import xlwings as xw
import os

w_Book = xw.Book()
dirname = os.getcwd()
save_folder_path =os.path.join(dirname,'保存先')
w_year =w_Book.sheets['売り上げ記入用'].range('F2').options(empty =2020).value
year_dict = {'{}年.xlsx'.format(w_year): os.path.join(save_folder_path,'{}年保存先.xlsx'.format(w_year))}
w_path= os.path.join(dirname,'書き込み用エクセルファイル.xlsm')
s_path= os.path.join(year_dict['{}年.xlsx'.format(w_year)])
xw1 = xw.Book(w_path)
xw2 = xw.Book(s_path)
sheet1 = xw1.sheets['売り上げ記入用']

def day_sum():
    def osi_sum():
        o_sum = sum([sheet1['B{}'.format(n1)].value * sheet1['C{}'.format(n1)].value for n1 in range(1, 5, 1)])
        return o_sum
    def water_sum():
        w_sum = sum([sheet1['B{}'.format(n2)].value * sheet1['C{}'.format(n2)].value for n2 in range(5, 10, 1)])
        return w_sum
    def ice_sum():
        i_sum = sum([sheet1['B{}'.format(n3)].value * sheet1['C{}'.format(n3)].value for n3 in range(10, 17, 1)])
        return i_sum
    wet = osi_sum()
    water = water_sum()
    ice = ice_sum()
    d_sum = wet+water+ice
    return d_sum
def month_sum():
    for m in range(1,13,1):
        sheet2 = xw2.sheets['{}月'.format(m)]
        m_sum= sum([sheet2['E{}'.format(num)].value for num in (2,32,1)])
        return m_sum
def year_sum():
    y_sheet = xw2.sheets['年'].range('B2:B13').value[0:][0:]
    y_sum = sum(y_sheet)
    return y_sum
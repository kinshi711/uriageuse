import os
from xlwings import Range,Book
import syuya
day_sum = syuya.day_sum()#日合計の関数を取得
month_sum = syuya.month_sum()#月合計の関数を取得
year_sum = syuya.year_sum()#年合計の関数を取得
wet_towel = syuya.osibori#おしぼりのインスタンス化されている関数を取得
water = syuya.water#水のインスタンス化されている関数を取得
ice = syuya.ice#氷のインスタンス化されている関数を取得
na = syuya.na

def myfunction():
    wb = Book.caller()
    Range('D1').value = "Call Python!"


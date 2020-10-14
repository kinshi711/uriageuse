import os
from xlwings import Range,Book

def myfunction():
    wb = Book.caller()
    Range('D1').value = "Call Python!"

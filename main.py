import os
from xlwings import Range,Book

def myfunction():
    wb = Book.caller()
    Range('A1').value = "Call Python!"
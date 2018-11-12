import openpyxl as xl
import requests as req
import datetime as dt
import time


class date:
    def __init__(self, loc):
        self.loc = loc
    def printdate(self):
        self.loc = time.time()
#class cols:
 #   def __init_(self, colA, colB, colC, colD):
  #      self.locA = locA
   #     self.locB = locB
    #    self.locC = locC
     #   self.locD = locD


wb = xl.load_workbook(filename = 'testing.xlsx')
ws = wb.active
#a = date(ws["H3"])
ws['H3'] = '123'
#date.printdate(a)






wb.save

import openpyxl as xl
import requests as req
import datetime as dt
import time as t

wb = xl.load_workbook(r'C:\Users\aaron.saliba\source\repos\test\testing.xlsx')
print(wb.sheetnames)



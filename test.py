import openpyxl as xl
import requests as req
import datetime as dt


class active_sheet:
    def __init__(self, work_book_loc, work_sheet):
        self.work_book_loc = work_book_loc
        self.work_sheet = work_sheet
    def active_testing_sheet(self):
        wb = xl.load_workbook(self.work_book_loc)
        ws = wb[self.work_sheet]
        

class c_date:
    def __init__(self, work_book_loc, work_sheet, cell):
        self.work_book_loc = work_book_loc
        self.work_sheet = work_sheet
        self.cell = cell
    def today_in_cell(self):
        wb = xl.load_workbook(self.work_book_loc)
        wb[self.work_sheet]
        dt.datetime.now()

a = active_sheet(r'C:\Users\aaron.saliba\source\repos\test\testing.xlsx',"Table")
a.active_testing_sheet()
b = c_date(r'C:\Users\aaron.saliba\source\repos\test\testing.xlsx', "Table", "G4")
b.today_in_cell()
c = active_workbook(r'C:\Users\aaron.saliba\source\repos\test\testing.xlsx')
aw = xl.
active_workbook.("testing.xlsx")

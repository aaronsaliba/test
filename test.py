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

def active_workbook(self):
        wb = xl.workbook.Workbook.active

ws = active_sheet("testing.xlsx","Table")
ws.active_testing_sheet()
date = c_date("testing.xlsx", "Table", "G4")
date.today_in_cell()
wb = active_workbook

import openpyxl as xl
import openpyxl.styles as xlst
import requests as req
import datetime as dt
import json
import excel_writer_reader as ex
import coinmarketcap_data_parser as cmc

# Timestamping
ex.today_in_cell("testing.xlsx", "Table", "G4")

# Filling in data in top-part
reading_set = ["A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A15", "A16", "A17"]
printing_set1 = ["B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17"]
printing_set2 = ["C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15", "C16", "C17"]

for x, y, z in zip(reading_set, printing_set1, printing_set2):
    a = ex.cell_reader("testing.xlsx", "Table", x)
    b,c = a.split("(")
    d = c.replace(")","")
    e = cmc.cmc_price(d,"EUR")
    f = cmc.cmc_7d(d,"EUR")/100
    ex.string_in_cell("testing.xlsx","Table", y, e)
    ex.string_in_cell("testing.xlsx","Table", z, f)


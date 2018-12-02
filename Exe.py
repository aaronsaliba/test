import openpyxl as xl
import openpyxl.styles as xlst
import requests as req
import datetime as dt
import json
import excel_writer_reader as ex
import coinmarketcap_data_parser as cmc

# Problem: cannot set the variable as the name
a = ex.cell_reader("testing.xlsx","Table","A4")
cmc.cmc_price(a ,"EUR")



# ex.cell_reader("testing.xlsx","Table","A5")
# c3 = ex.cell_reader("testing.xlsx","Table","A6")
# c4 = ex.cell_reader("testing.xlsx","Table","A7")
# c5 = ex.cell_reader("testing.xlsx","Table","A8")
# c6 = ex.cell_reader("testing.xlsx","Table","A9")
# c7 = ex.cell_reader("testing.xlsx","Table","A10")
# c8 = ex.cell_reader("testing.xlsx","Table","A11")
# c9 = ex.cell_reader("testing.xlsx","Table","A12")
# c10 = ex.cell_reader("testing.xlsx","Table","A13")
# c11 = ex.cell_reader("testing.xlsx","Table","A14")
# c12 = ex.cell_reader("testing.xlsx","Table","A15")
# c13 = ex.cell_reader("testing.xlsx","Table","A16")
# c14 = ex.cell_reader("testing.xlsx","Table","A17")




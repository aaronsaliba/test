import openpyxl as xl
import openpyxl.styles as xlst
import requests as req
import datetime as dt
import json
import excel_writer_reader as ex
import coinmarketcap_data_parser as cmc

a1 = ex.cell_reader("testing.xlsx","Table","A4")
b1,c1 = a1.split("(")
d1 = c1.replace(")","")
e1 = cmc.cmc_price(d1,"EUR")
f1 = cmc.cmc_7d(d1,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B4", e1)
ex.string_in_cell("testing.xlsx","Table","C4", f1)

a2 = ex.cell_reader("testing.xlsx","Table","A5")
b2,c2 = a2.split("(")
d2 = c2.replace(")","")
e2 = cmc.cmc_price(d2,"EUR")
f2 = cmc.cmc_7d(d2,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B5", e2)
ex.string_in_cell("testing.xlsx","Table","C5", f2)

a3 = ex.cell_reader("testing.xlsx","Table","A6")
b3,c3 = a3.split("(")
d3 = c3.replace(")","")
e3 = cmc.cmc_price(d3 ,"EUR")
f3 = cmc.cmc_7d(d3,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B6", e3)
ex.string_in_cell("testing.xlsx","Table","C6", f3)

a4 = ex.cell_reader("testing.xlsx","Table","A7")
b4,c4 = a4.split("(")
d4 = c4.replace(")","")
e4 = cmc.cmc_price(d4 ,"EUR")
f4 = cmc.cmc_7d(d4,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B7", e4)
ex.string_in_cell("testing.xlsx","Table","C7", f4)

a5 = ex.cell_reader("testing.xlsx","Table","A8")
b5,c5 = a5.split("(")
d5 = c5.replace(")","")
e5 = cmc.cmc_price(d5 ,"EUR")
f5 = cmc.cmc_7d(d5,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B8", e5)
ex.string_in_cell("testing.xlsx","Table","C8", f5)

a6 = ex.cell_reader("testing.xlsx","Table","A9")
b6,c6 = a6.split("(")
d6 = c6.replace(")","")
e6 = cmc.cmc_price(d6 ,"EUR")
f6 = cmc.cmc_7d(d6,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B9", e6)
ex.string_in_cell("testing.xlsx","Table","C9", f6)

a7 = ex.cell_reader("testing.xlsx","Table","A10")
b7,c7 = a7.split("(")
d7 = c7.replace(")","")
e7 = cmc.cmc_price(d7 ,"EUR")
f7 = cmc.cmc_7d(d7,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B10", e7)
ex.string_in_cell("testing.xlsx","Table","C10", f7)

a8 = ex.cell_reader("testing.xlsx","Table","A11")
b8,c8 = a8.split("(")
d8 = c8.replace(")","")
e8 = cmc.cmc_price(d8 ,"EUR")
f8 = cmc.cmc_7d(d8,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B11", e8)
ex.string_in_cell("testing.xlsx","Table","C11", f8)

a9 = ex.cell_reader("testing.xlsx","Table","A12")
b9,c9 = a9.split("(")
d9 = c9.replace(")","")
e9 = cmc.cmc_price(d9 ,"EUR")
f9 = cmc.cmc_7d(d9,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B12", e9)
ex.string_in_cell("testing.xlsx","Table","C12", f9)

a10 = ex.cell_reader("testing.xlsx","Table","A13")
b10,c10 = a10.split("(")
d10 = c10.replace(")","")
e10 = cmc.cmc_price(d10 ,"EUR")
f10 = cmc.cmc_7d(d10,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B13", e10)
ex.string_in_cell("testing.xlsx","Table","C13", f10)

a11 = ex.cell_reader("testing.xlsx","Table","A14")
b11,c11 = a11.split("(")
d11 = c11.replace(")","")
e11 = cmc.cmc_price(d11 ,"EUR")
f11 = cmc.cmc_7d(d11,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B14", e11)
ex.string_in_cell("testing.xlsx","Table","C14", f11)

a12 = ex.cell_reader("testing.xlsx","Table","A15")
b12,c12 = a12.split("(")
d12 = c12.replace(")","")
e12 = cmc.cmc_price(d12 ,"EUR")
f12 = cmc.cmc_7d(d12,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B15", e12)
ex.string_in_cell("testing.xlsx","Table","C15", f12)

a13 = ex.cell_reader("testing.xlsx","Table","A16")
b13,c13 = a13.split("(")
d13 = c13.replace(")","")
e13 = cmc.cmc_price(d13 ,"EUR")
f13 = cmc.cmc_7d(d13,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B16", e13)
ex.string_in_cell("testing.xlsx","Table","C16", f13)

a14 = ex.cell_reader("testing.xlsx","Table","A17")
b14,c14 = a14.split("(")
d14 = c14.replace(")","")
e14 = cmc.cmc_price(d14 ,"EUR")
f14 = cmc.cmc_7d(d14,"EUR")/100
ex.string_in_cell("testing.xlsx","Table","B17", e14)
ex.string_in_cell("testing.xlsx","Table","C17", f14)

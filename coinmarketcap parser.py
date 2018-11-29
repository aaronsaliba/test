import openpyxl as xl
import openpyxl.styles as xlst
import requests as req
import datetime as dt
import excel_writer as writer
import json

def cmc_data ():
    cmc = req.get("https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest", headers={"X-CMC_PRO_API_KEY":"9d1daf10-3b0e-4004-8abe-957e955940a5"})
    print (cmc.json)

cmc_data()




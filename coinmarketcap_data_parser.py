import openpyxl as xl
import openpyxl.styles as xlst
import requests as req
import datetime as dt
import excel_writer_reader as ex
import json

def cmc_price (name, currency):
    """[Extracts price from coinmarketcap]
    
    Arguments:
        name {[string]} -- [name of coin]
        currency {[string]} -- [currency of the price desired]
    """

    if currency == "USD":
        r = req.get("https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?convert=USD", headers={"X-CMC_PRO_API_KEY":"9d1daf10-3b0e-4004-8abe-957e955940a5"})
    elif currency == "EUR":
        r = req.get("https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?convert=EUR", headers={"X-CMC_PRO_API_KEY":"9d1daf10-3b0e-4004-8abe-957e955940a5"})
    else:
        print ("Error: Currency not available")
    if r.status_code == 200:
        cmc = json.loads(r.text)
        cmc_data = cmc['data']
        for price in cmc_data:
            if name == price["name"]:
                try: 
                    print(price['quote'][currency]["price"])
                except:
                    print("error")
    else:
        print("Error: Data cannot be accessed")

    



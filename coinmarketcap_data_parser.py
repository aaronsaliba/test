import openpyxl as xl
import openpyxl.styles as xlst
import requests as req
import datetime as dt
import excel_writer_reader as ex
import json

def cmc_price (symbol, currency):
    """[Extracts price from coinmarketcap]
    
    Arguments:
        symbol {[string]} -- [symbol of coin]
        currency {[string]} -- [currency of the price desired]
    """
    url = f"https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol={symbol}&convert={currency}"

    r = req.get(url, headers={"X-CMC_PRO_API_KEY":"9d1daf10-3b0e-4004-8abe-957e955940a5"})

    if r.status_code == 200:
        cmc = json.loads(r.text)
        return cmc['data'][ symbol ][ 'quote' ][ currency ][ 'price' ]                   
    else:
        return "Error: Data cannot be accessed"

def cmc_7d (symbol, currency):
    """[Extracts 7d% change from coinmarketcap]
    
    Arguments:
        symbol {[string]} -- [symbol of coin]
        currency {[string]} -- [currency of the price desired]
    """
    url = f"https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol={symbol}&convert={currency}"

    r = req.get(url, headers={"X-CMC_PRO_API_KEY":"9d1daf10-3b0e-4004-8abe-957e955940a5"})

    if r.status_code == 200:
        cmc = json.loads(r.text)
        return cmc['data'][ symbol ][ 'quote' ][ currency ][ 'percent_change_7d' ]                   
    else:
        return "Error: Data cannot be accessed"

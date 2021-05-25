import requests
from bs4 import BeautifulSoup as bs
import time
import pandas as pd

def one_page_list(sosok, page):
    STOCKLIST_URL = "https://finance.naver.com/sise/sise_market_sum.nhn?sosok={}&page={}".format(sosok, page) #주소설정
    STOCK_NAME_LIST = []
    STOCK_LIST = []

    response = requests.get(STOCKLIST_URL)
    soup = bs(response.content.decode('euc-kr', 'replace'), 'html.parser')

    for tr in soup.findAll('tr'):
        stockName = tr.findAll('a', attrs={'class', 'tltle'})
        if stockName is None or stockName == []:
            pass
        else:
            stockName = stockName[0].contents[-1]
            STOCK_NAME_LIST.append(stockName)
            
    for i in range(len(STOCK_NAME_LIST)):
        stockInfo = [STOCK_NAME_LIST[i]]
        STOCK_LIST.append(stockInfo)

    return pd.DataFrame(STOCK_LIST)


def all_page_list():
    print("Start crwaling...")

    FINAL_LIST = []

    for sosok in range(2):
        for page in range(33):
            one_page_data = one_page_list(sosok, page+1)
            
            if one_page_data is None:
                break
            
            FINAL_LIST.append(one_page_data)
            time.sleep(3)

    print("Finished")
    return pd.concat(FINAL_LIST)

ALL_STOCK_LIST = all_page_list()
ALL_STOCK_LIST.to_csv('./stocks.csv', header=False, index=False, sep=',', na_rep='NaN') 


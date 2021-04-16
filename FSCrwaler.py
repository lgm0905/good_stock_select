from re import L
import sys, os
from numpy.core.numeric import NaN
import requests
from bs4 import BeautifulSoup as bs
import time
import pandas as pd
import csv
from selenium import webdriver
import OpenDartReader


def GetStockDataFrame():
    stock_df = pd.read_excel('data.xlsx', usecols=['단축코드', '한글 종목명'])
    return stock_df


def SelectQuarter(year, month, rng) :
    qtr_li = []
    
    for y in range(year-2000-rng, year-2000) :
        for m in range(1,13) :
            if m % 3 is 0 :
                qtr = str(y) + "년 " + str(m) + "월"
                qtr_li.append(qtr)

    for m in range(1,month) :
        if m % 3 is 0 :
            qtr = str(year-2000) + "년 " + str(m) + "월"
            qtr_li.append(qtr)

    return qtr_li


def FSCrawler(stocks):
    # driver = webdriver.PhantomJS(os.path.join('./phantomjs/phantomjs-2.1.1-linux-x86_64/bin/phantomjs'))
    driver = webdriver.PhantomJS(os.path.join('./phantomjs/phantomjs-2.1.1-macosx/bin/phantomjs'))
    
    driver.implicitly_wait(3)

    qrt_3year = SelectQuarter(2021,4,3)
    qrt_1year = SelectQuarter(2021,4,1)

    fs_dict = {'종목명' : [], '종목코드' : [], '시가총액' : [], 'ROE_3' : [], 'ROE_1' : [], 'OM_3' : [], 'OM_1' : []}

    # temp_cnt = 0

    for idx, row in stocks.iterrows():
        # if temp_cnt is 1 :
        #     break

        # temp_cnt += 1

        code = row['단축코드']; name = row['한글 종목명']
        print(str(idx+1) + ' ' + name)

        ROE_3 = 0; ROE_1 = 0; OM_3 = 0; OM_1 = 0
        cnt_3y = 0; cnt_1y = 0

        # ROE, 영업이익률 크롤링
        driver.get('https://stockplus.com/m/stocks/KOREA-A{}/analysis'.format(code))
        time.sleep(1.5)
        raw = driver.page_source
        soup = bs(raw, 'html.parser')
        table = soup.select_one('.type02 tbody')

        if table is None : 
            print("정보가 없습니다.")
            continue

        fs_dict['종목명'].append(name)
        fs_dict['종목코드'].append(code)

        for quarter, sales, operatingProfit, netIncome, operatingMargin, netProfitMargin, PER, PBR, ROE in zip(table.select('tr')[0].select('th'), table.select('tr')[1].select('td'), table.select('tr')[2].select('td'), table.select('tr')[3].select('td'), table.select('tr')[4].select('td'), table.select('tr')[5].select('td'), table.select('tr')[6].select('td'), table.select('tr')[7].select('td'), table.select('tr')[8].select('td')):
            if 'E' in quarter.text:
                break
            
            if quarter.text.split('월')[0]+'월' in qrt_3year : 
                if ROE.text != '-' :
                    ROE_3 += float(ROE.text.replace(',',''))
                if operatingMargin.text != '-' :  
                    OM_3 += float(operatingMargin.text.replace(',',''))
                cnt_3y += 1

            if quarter.text.split('월')[0]+'월' in qrt_1year : 
                if ROE.text != '-' :
                    ROE_1 += float(ROE.text.replace(',',''))
                if operatingMargin.text != '-' :
                    OM_1 += float(operatingMargin.text.replace(',',''))
                cnt_1y += 1
        
        if cnt_3y is not 0 : 
            fs_dict['ROE_3'].append(round(ROE_3/cnt_3y, 3))
            fs_dict['OM_3'].append(round(OM_3/cnt_3y, 3))
        else :
            fs_dict['ROE_3'].append(0)
            fs_dict['OM_3'].append(0)

        if cnt_1y is not 0 : 
            fs_dict['ROE_1'].append(round(ROE_1/cnt_1y,3))
            fs_dict['OM_1'].append(round(OM_1/cnt_1y,3))
        else :
            fs_dict['ROE_1'].append(0)
            fs_dict['OM_1'].append(0)


        # 시가총액 크롤링
        driver.get('https://stockplus.com/m/stocks/KOREA-A{}'.format(code))
        time.sleep(1.5)
        raw = driver.page_source
        soup = bs(raw, 'html.parser')
        table = soup.find('div', class_='ftHiLowB pt0')
        capital = table.select('tr')[4].select('td')[0].select('span')[0].text
        capital = int(capital.replace(',',''))
        
        fs_dict['시가총액'].append(capital)

    driver.quit()
    return fs_dict


def SetRank(fs_dict):
    df = pd.DataFrame(data=fs_dict)
    df['ROE_3_RANK'] = df['ROE_3'].rank(method='max', ascending=False)
    df['OM_3_RANK'] = df['OM_3'].rank(method='max', ascending=False)
    df['ROE_1_RANK'] = df['ROE_1'].rank(method='max', ascending=False)
    df['OM_1_RANK'] = df['OM_1'].rank(method='max', ascending=False)
    df['FINAL_3_RANK'] = (df['OM_3_RANK'] + df['ROE_3_RANK']).rank(method='min', ascending=True)
    df['FINAL_1_RANK'] = (df['OM_1_RANK'] + df['ROE_1_RANK']).rank(method='min', ascending=True)

    return df


def ChangeColName(df) :
    df.rename(columns={ 'ROE_3':            '지난 3년 ROE', 
                        'ROE_1':            '최근 ROE', 
                        'OM_3' :            '지난 3년 영업이익률', 
                        'OM_1' :            '최근 영업이익률',
                        'ROE_3_RANK' :      '지난 3년 ROE 순위',
                        'OM_3_RANK' :       '지난 3년 영업이익률 순위',
                        'ROE_1_RANK' :      '최근 ROE 순위',
                        'OM_1_RANK' :       '최근 영업이익률 순위',
                        'FINAL_3_RANK' :    '지난 3년 최종순위',
                        'FINAL_1_RANK' :    '최근 최종순위',
                        }, inplace=True)

    return df


if __name__ == '__main__':
    stocks = GetStockDataFrame()
    fs_dict = FSCrawler(stocks)
    df = SetRank(fs_dict)
    df = ChangeColName(df)
    df.to_excel('result.xlsx')
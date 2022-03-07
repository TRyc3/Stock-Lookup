from asyncio.windows_events import INFINITE
from cgitb import reset
from re import X
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import PySimpleGUI as sg


#window layout
layout = [
        #take stock symbol as input
        [sg.Text("Symbol"), 
        sg.InputText(key = "Symbol"),

        #select period
        sg.Listbox(values = [
            "Six Months",
            "YTD",
            "One Year",
            "Five Years",
            "Max",],
            key = "Period",
            size = (10,5))],

        #search button starts look up process
        #done closes program
        [sg.Button("Search"),
        sg.Button("Done")]
    ]

window = sg.Window("Stock Look up", layout)
while True:
    event,values = window.read()
    if event == sg.WIN_CLOSED or event == "Done":
        break

    if event == "Search":
        driver = webdriver.Chrome()

        symbol = values["Symbol"]

        link = "https://finance.yahoo.com/quote/" + symbol + "/history" 

        driver.get(link)

        soup = BeautifulSoup(driver.page_source, "html.parser")

        #open period selection drop down menu
        print ("loading")
        element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="M(0) O(n):f D(ib) Bd(0) dateRangeBtn O(n):f Pos(r)"]')))
        element.click()
        
        driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, Keys.PAGE_DOWN)
        driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, Keys.DOWN)


        #select period based on user input
        if values["Period"][0] == "Six Months":
            element = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-value="6_M"]')))
            element.click()
        elif values["Period"][0] == "YTD":
            element = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-value="YTD"]')))
            element.click()
        elif values["Period"][0] == "One Year":
            element = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-value="1_Y"]')))
            element.click()
        elif values["Period"][0] == "Five Years":
            element = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-value="5_Y"]')))
            element.click()
        elif values["Period"][0] == "MAX":
            element = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@data-value="MAX"]')))
            element.click()

        print ("apply")
        #click apply button to change results based on desired time frame
        element = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//button/span[contains(text(),"Apply")]')))
        element.click()
        # driver.find_element(By.XPATH, '//button/span[contains(text(),"Apply")]').click()


        print ("scroll")
        #scroll to end of results
        table_end = soup.find_all("div", {"class" : "Mstart(30px) Pt(10px)"})
        while len(table_end) > 0:
            driver.find_element(By.CSS_SELECTOR, "body").send_keys(Keys.CONTROL, Keys.END)
            soup = BeautifulSoup(driver.page_source, "html.parser")
            table_end = soup.find_all("div", {"class" : "Mstart(30px) Pt(10px)"})
        driver.close()

        #locate and collect table information
        table = soup.find("table", {"class" : "W(100%) M(0)"})
        results = table.find_all("tr", {"class" : "BdT Bdc($seperatorColor) Ta(end) Fz(s) Whs(nw)"})

        header = ['DATE', 'OPEN', 'HIGH', 'LOW', 'CLOSE', 'ADJ CLOSE', 'VOLUME']

        #parse data and remove unneeded rows
        data = []
        for i in results:
            col = 0
            day_data = i.find_all("span")
            info = []
            for j in day_data:
                res = j.contents
                res = str(res).replace("['",'')
                res = res.replace("']",'')
                res = res.replace(',','')
                info.append(res)
                col += 1
            if col == 7:
                data.append(info)

        #create dataframe
        df = pd.DataFrame(data = data, columns = header)
        df.astype({'OPEN' : float, 'HIGH' : float, 'LOW' : float, 'ADJ CLOSE' : float, 'VOLUME' : float})
        df.set_index("DATE", inplace = True)

        #change columns to lists for easier computation
        volume = df['VOLUME'].to_list()
        close = df['CLOSE'].to_list()
        open = df['OPEN'].to_list()
        high = df['HIGH'].to_list()
        low = df['LOW'].to_list()
        volume.reverse()
        close.reverse()
        open.reverse()
        high.reverse()
        low.reverse()


        #calculate on-balance volume
        obv = [int(volume[0])]
        for i in range (1, len(volume)):
            if close[i] > close[i - 1]:
                obv.append(obv[i - 1] + int(volume[i]))
            elif close[i] < close[i - 1]:
                obv.append(obv[i - 1] - int(volume[i]))
            else:
                obv.append(obv[i - 1])
        obv.reverse()

        #calculate accumulation-distribution line
        ad = [(((float(close[0]) - float(low[0])) - (float(high[0]) - float(close[0]))) / (float(high[0]) - float(low[0]))) * float(volume[0])]
        for i in  range(1, len(close)):
            cmfv = (((float(close[i]) - float(low[i])) - (float(high[i]) - float(close[i]))) / (float(high[i]) - float(low[i]))) * float(volume[i])
            ad.append(ad[i - 1] + cmfv)
        ad.reverse()

        #calculate Aroon oscilator
        aroon = []
        for i in range (len(high)):
            period_high = 0
            max = 0
            period_low = 0
            min = 0
            if i >= 25:
                for j in range(25):
                    if float(high[i - j]) > max:
                        max = float(high[i - j])
                        period_high = j
                    if float(low[i - j]) < min:
                        min = float(low[i - j])
                        period_low = j
            else:
                for j in range(i):
                    if float(high[i - j]) > max:
                        max = float(high[i - j])
                        period_high = j
                    if float(low[i - j]) < min:
                        min = float(low[i - j])
                        period_low = j

            aroon_up = 4 * (25 - period_high)
            aroon_down = 4 * (25 - period_low)
            aroon_osc = aroon_up - aroon_down
            aroon.append(aroon_osc)
        aroon.reverse()


        #calculate relative strength index
        rsi = []
        for i in range(len(close)):
            profit = 0
            loss = 0
            average_profit = 0
            average_loss = 0
            if i >= 14:
                for j in range(14):
                    if float(close[i - j]) > float(open[i - j]):
                        profit += float(close[i - j]) - float(open[i - j])
                    elif float(close[i - j]) < float(open[i - j]):
                        loss += float(open[i - j]) - float(close[i - j])
                average_profit = profit / 14
                average_loss = loss / 14
            else:
                for j in range(i):
                    if float(close[i - j]) > float(open[i - j]):
                        profit += float(close[i - j]) - float(open[i - j])
                    elif float(close[i - j]) < float(open[i - j]):
                        loss += float(open[i - j]) - float(close[i - j])
                average_profit = profit / (i + 1)
                average_loss = loss / (i + 1)
            if average_loss == 0:
                rs = INFINITE
            else:
                rs = average_profit / average_loss
            rsi.append(100 - (100 / (1 + rs)))
        rsi.reverse()


        #calculate moving average convergence divergence
        macd = []
        signal = []
        for i in range(len(close)):
            ema_nine = 0
            ema_twelve = 0
            ema_twentysix = 0
            if i >= 26:
                for j in range(26, 0, -1):
                    ema_twentysix = (float(close[i - j]) * (2/27)) + (ema_twentysix * (2/27))
                    if j < 12:
                        ema_twelve = (float(close[i - j]) * (2/13)) + (ema_twelve * (2/13))
                    if j < 9:
                        ema_nine = (float(close[i - j]) * (2/10)) + (ema_nine * (2/10))
            else:
                for j in range(j, 0, -1):
                    ema_twentysix = (float(close[i - j]) * (2/27)) + (ema_twentysix * (2/27))
                    if j < 12:
                        ema_twelve = (float(close[i - j]) * (2/13)) + (ema_twelve * (2/13))
                    if j < 9:
                        ema_nine = (float(close[i - j]) * (2/10)) + (ema_nine * (2/10))
            macd.append(ema_twelve - ema_twentysix)
            signal.append(ema_nine)
        macd.reverse()
        signal.reverse


        #add analysis to dataframe
        df.insert(6, 'OBV', obv, True)
        df.insert(7, 'A/D', ad, True)
        df.insert(8, 'AROON', aroon, True)
        df.insert(9, 'RSI', rsi, True)
        df.insert(10, 'MACD', macd, True)
        df.insert(11, 'SIGNAL', signal, True)

        print (df)

        # date_tick = [df["DATE"][0]]
        # date_tick_location = [0]
        # for i in range(1,len(df["DATE"])):
        #     if df["DATE"][i][2] != df["DATE"][i - 1][2] and df["DATE"][i][1] != df["DATE"][i - 1][1]:
        #         date_tick.append(df["DATE"][i])
        #         date_tick_location.append(i)

        #visualization
        # plt.plot(df["DATE"], df["OBV"])   
        # plt.xlabel('Date')
        # plt.ylabel('On-Balance Volume')
        # plt.axes().set_xticklabels(date_tick, reset = True)
        # plt.show()

        plt.subplots(figsize = (15, 20))
        plt.xticks(rotation = 45)
        res = sns.lineplot(data = df["OBV"])
        res.set_xlabel("Dates", fontsize = 20)
        for index, label in enumerate(res.get_xticklabels()):
            if values["Period"][0] == "Five Years" or values["Period"][0] == "MAX":
                if index % 100 == 0:
                    label.set_visible(True)
                else:
                    label.set_visible(False)
            elif values["Period"][0] == "One Year" or values["Period"][0] == "YTD" or values["Period"][0] == "Six Months":
                if index % 25 == 0:
                    label.set_visible(True)
                else:
                    label.set_visible(False)
        # res.set_xticklabels(res.get_xmajorticklabels(), fontsize = 6)
        # res.set_yticklabels(res.get_ymajorticklabels(), fontsize = 6)
        sns.set(font_scale=1)
        plt.show()
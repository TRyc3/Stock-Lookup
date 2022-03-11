from asyncio.windows_events import INFINITE
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from itertools import islice
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
        
        driver.execute_script("window.scrollTo(0, 500)")


        error = soup.find_all("span", {"class" : "D(b) Ta(c) W(100%) Fz(m) C($tertiaryColor) Mb(10px) Fw(500) Ell"})
        if len(error) > 0:
            print ("ERROR: Symbol not valid")
            driver.close()
            sg.Print("An error happened. Symbol Invalid")
            sg.popup_error('AN EXCEPTION OCCURRED!')
            break

        #open period selection drop down menu
        element = WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.XPATH, '//div[@class="M(0) O(n):f D(ib) Bd(0) dateRangeBtn O(n):f Pos(r)"]')))
        element.click()

        #select period based on user input
        if len(values["Period"]) > 0:
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

            #click apply button to change results based on desired time frame
            element = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, '//button/span[contains(text(),"Apply")]')))
            element.click()


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
        date = []
        open = []
        high = []
        low = []
        close = []
        adj_close = []
        volume = []
        for i in range(len(data)):
            date.append(data[i][0])
            open.append(float(data[i][1]))
            high.append(float(data[i][2]))
            low.append(float(data[i][3]))
            close.append(float(data[i][4]))
            adj_close.append(float(data[i][5]))
            volume.append(float(data[i][6]))

        date.reverse()
        open.reverse()
        high.reverse()
        low.reverse()
        close.reverse()
        adj_close.reverse()
        volume.reverse()


        df = pd.DataFrame(data = date, columns = ["DATE"])
        df.insert(1, 'OPEN', open, True)
        df.insert(2, 'HIGH', high, True)
        df.insert(3, 'LOW', low, True)
        df.insert(4, 'CLOSE', close, True)
        df.insert(5, 'ADJ', adj_close, True)
        df.insert(6, 'VOLUME', volume, True)
        df.set_index("DATE", inplace = True)

        #calculate on-balance volume
        obv = [int(volume[0])]
        for i in range (1, len(volume)):
            if close[i] > close[i - 1]:
                obv.append(obv[i - 1] + int(volume[i]))
            elif close[i] < close[i - 1]:
                obv.append(obv[i - 1] - int(volume[i]))
            else:
                obv.append(obv[i - 1])

        #calculate accumulation-distribution line
        ad = [(((float(close[0]) - float(low[0])) - (float(high[0]) - float(close[0]))) / (float(high[0]) - float(low[0]))) * float(volume[0])]
        for i in  range(1, len(close)):
            cmfv = (((float(close[i]) - float(low[i])) - (float(high[i]) - float(close[i]))) / (float(high[i]) - float(low[i]))) * float(volume[i])
            ad.append(ad[i - 1] + cmfv)

        #calculate Aroon oscilator
        aroon = []
        for i in range (len(high)):
            period_high = 0
            max = 0
            period_low = 0
            min = 0
            if i >= 25:
                for j in range(25):
                    if high[i - j] > max:
                        max = high[i - j]
                        period_high = j
                    if low[i - j] < min:
                        min = low[i - j]
                        period_low = j
            else:
                for j in range(i):
                    if high[i - j] > max:
                        max = high[i - j]
                        period_high = j
                    if float(low[i - j]) < min:
                        min = low[i - j]
                        period_low = j

            aroon_up = 4 * (25 - period_high)
            aroon_down = 4 * (25 - period_low)
            aroon_osc = aroon_up - aroon_down
            aroon.append(aroon_osc)


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


        #calculate moving average convergence divergence
        macd = []
        signal = []
        ma_twentysix = []
        ma_twelve = []
        ma_nine = []
        for i in range(len(close)):
            ema_nine = 0
            ema_twelve = 0
            ema_twentysix = 0
            if i >= 26:
                ema_twentysix = (float(close[i]) * (2/27)) + (ma_twentysix[i - 1] * (1 - (2/27)))
                ema_twelve = (float(close[i]) * (2/13)) + (ma_twelve[i - 1] * (1 - (2/13)))
                ema_nine = (float(close[i]) * (2/10)) + (ma_nine[i - 1] * (1 - (2/10)))
            else:
                for j in range(i):
                    ema_twentysix += float(close[i - j])
                    if j < 12:
                        ema_twelve += float(close[i - j])
                    if j < 9:
                        ema_nine += float(close[i - j])

                ema_twentysix = ema_twentysix / (i + 1)
                if i >= 12:
                    ema_twelve = ema_twelve / 12
                else:
                    ema_twelve = ema_twelve / (i + 1)
                if i >= 9:
                    ema_nine = ema_nine / 9
                else:
                    ema_nine = ema_nine / (i + 1)

        
            macd.append(ema_twelve - ema_twentysix)
            signal.append(ema_nine - ema_twentysix)
            ma_twentysix.append(ema_twentysix)
            ma_twelve.append(ema_twelve)
            ma_nine.append(ema_nine)

        #add analysis to dataframe
        df.insert(6, 'OBV', obv, True)
        df.insert(7, 'A/D', ad, True)
        df.insert(8, 'AROON', aroon, True)
        df.insert(9, 'RSI', rsi, True)
        df.insert(10, 'MACD', macd, True)
        df.insert(11, 'SIGNAL', signal, True)

        print (df.dtypes)
        print(df)

        #Second menu
        layout = [
            [sg.Text(symbol.upper())],

            #Buttons for graphs
            [sg.Button("OBV"),
            sg.Button("A/D"),
            sg.Button("AROON"),
            sg.Button("RSI"),
            sg.Button("MACD"),
            sg.Button("Done")]
        ]
        graph_window = sg.Window("Stock Look up", layout)

        while True:
            graph_event, graph_value = graph_window.read()

            #visualization
            fig, axes = plt.subplots(2, 1, figsize = (20, 10))
            if graph_event == "OBV":
                res = sns.lineplot(ax = axes[1], data = df["OBV"], legend = 'auto')
            elif graph_event == "A/D":
                res = sns.lineplot(ax = axes[1], data = df["A/D"], legend = 'auto')
            elif graph_event == "AROON":
                res = sns.lineplot(ax = axes[1], data = df["AROON"], legend = 'auto')
            elif graph_event == "RSI":
                res = sns.lineplot(ax = axes[1], data = df["RSI"], legend = 'auto')
            elif graph_event == "MACD":
                res = sns.lineplot(ax = axes[1], data = df["MACD"], legend = 'auto')
                res = sns.lineplot(ax = axes[1], data = df["SIGNAL"], legend = 'auto')

            for index, label in enumerate(res.get_xticklabels()):
                if len(values["Period"]) > 0:    
                    if values["Period"][0] == "Five Years" or values["Period"][0] == "MAX":
                        if index % 100 == 0:
                            label.set_visible(True)
                        else:
                            label.set_visible(False)
                    else:
                        if index % 25 == 0:
                            label.set_visible(True)
                        else:
                            label.set_visible(False)
                else:
                    if index % 25 == 0:
                        label.set_visible(True)
                    else:
                        label.set_visible(False)
                

            res = sns.lineplot(ax = axes[0], data = df["CLOSE"], legend = 'auto')
            for index, label in enumerate(res.get_xticklabels()):
                if len(values["Period"]) > 0:    
                    if values["Period"][0] == "Five Years" or values["Period"][0] == "MAX":
                        if index % 100 == 0:
                            label.set_visible(True)
                        else:
                            label.set_visible(False)
                    else:
                        if index % 25 == 0:
                            label.set_visible(True)
                        else:
                            label.set_visible(False)
                else:
                    if index % 25 == 0:
                        label.set_visible(True)
                    else:
                        label.set_visible(False)
                
                    
            plt.show()

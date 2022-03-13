from asyncio.windows_events import INFINITE
from cProfile import label
from turtle import color
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from itertools import islice
import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import mplfinance as mpl
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
        window.close()
        break

    if event == "Search":
        symbol = values["Symbol"]
        if len(symbol) > 5 or len(symbol) == 0:
            print ("ERROR: Symbol not valid")
            sg.Print("An error happened. Symbol Invalid")
            sg.popup_error('AN EXCEPTION OCCURRED!')
            break  

        driver = webdriver.Chrome()                  

        link = "https://finance.yahoo.com/quote/" + symbol + "/history" 

        driver.get(link)

        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        driver.execute_script("window.scrollTo(0, 500)")

        table = soup.find("table", {"class" : "W(100%) M(0)"})
        results = table.find_all("tr", {"class" : "BdT Bdc($seperatorColor) Ta(end) Fz(s) Whs(nw)"})
        error = soup.find_all("span", {"class" : "D(b) Ta(c) W(100%) Fz(m) C($tertiaryColor) Mb(10px) Fw(500) Ell"})
        if len(error) > 0:
            print ("ERROR: Symbol not valid")
            driver.close()
            sg.Print("An error happened. Symbol Invalid")
            sg.popup_error('AN EXCEPTION OCCURRED!')
            break
        elif not len(results) > 0:
            print ("ERROR: Symbol not valid")
            driver.close()
            sg.Print("An error happened. Symbol Invalid")
            sg.popup_error('AN EXCEPTION OCCURRED!')

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
            datetime_object = datetime.strptime(data[i][0], '%b %d %Y')
            date.append(datetime_object)
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
        obv = [volume[0]]
        for i in range (1, len(volume)):
            if close[i] > close[i - 1]:
                obv.append(obv[i - 1] + volume[i])
            elif close[i] < close[i - 1]:
                obv.append(obv[i - 1] - volume[i])
            else:
                obv.append(obv[i - 1])

        #calculate accumulation-distribution line
        ad = [(((close[0] - low[0]) - (high[0] - close[0])) / (high[0] - low[0])) * volume[0]]
        for i in  range(1, len(close)):
            cmfv = (((close[i] - low[i]) - (high[i] - close[i])) / (high[i] - low[i])) * volume[i]
            ad.append(ad[i - 1] + cmfv)

        #calculate Aroon oscilator
        aroon_up = []
        aroon_down = []
        for i in range (len(high)):
            period_high = 0
            max = 0
            period_low = 0
            min = INFINITE
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
                    if low[i - j] < min:
                        min = low[i - j]
                        period_low = j

            aroon_up.append(4 * (25 - period_high))
            aroon_down.append(4 * (25 - period_low))


        #calculate relative strength index
        rsi = []
        average_gain = []
        average_loss = []
        for i in range(len(close)):
            profit = 0
            loss = 0
            current_profit = 0
            current_loss = 0
            if i >= 14:
                for j in range(1,13):
                    if close[i - j] > open[i - j]:
                        profit += (close[i - j] - open[i - j]) / open[i - j]
                    elif close[i - j] < open[i - j]:
                        loss += (open[i - j] - close[i - j]) / close[i - j]
                profit = (profit / 14) * 100
                loss = (loss / 14) * 100

                if close[i] > open[i]:
                    current_profit = (close[i] - open[i]) / open[i]
                    profit += current_profit
                elif close[i] < open[i]:
                    current_loss = (open[i] - close[i]) / close[i]
                    loss += current_loss

                profit = (profit / 14) * 100
                loss = (loss / 14) * 100
            else:
                for j in range(i):
                    if close[i - j] > open[i - j]:
                        profit += (close[i - j] - open[i - j]) / open[i - j]
                    elif close[i - j] < open[i - j]:
                        loss += (open[i - j] - close[i - j]) / close[i - j]
                profit = (profit / (i + 1)) * 100
                loss = (loss / (i + 1)) * 100
            if loss == 0:
                rs = INFINITE
            else:
                rs = profit / loss
                
            if i > 14:
                rsi.append(100 - (100 / (1 + ((average_gain[i - 1] * 13 ) + current_profit) / ((average_loss[i - 1] * 13) + current_loss))))
            else:
                rsi.append(100 - (100 / (1 + rs)))
            average_gain.append(profit)
            average_loss.append(loss)



        #calculate moving average convergence divergence
        macd = []
        signal = []
        ma_twentysix = []
        ma_twelve = []
        macd_ema = []
        for i in range(len(close)):
            ema_signal = 0
            ema_twelve = 0
            ema_twentysix = 0
            if i >= 26:
                ema_twentysix = (close[i] * (2/27)) + (ma_twentysix[i - 1] * (1 - (2/27)))
                ema_twelve = (close[i] * (2/13)) + (ma_twelve[i - 1] * (1 - (2/13)))
            else:
                for j in range(i):
                    ema_twentysix += close[i - j]
                    if j < 12:
                        ema_twelve += close[i - j]

                ema_twentysix = ema_twentysix / (i + 1)
                if i >= 12:
                    ema_twelve = ema_twelve / 12
                else:
                    ema_twelve = ema_twelve / (i + 1)

            macd.append(ema_twelve - ema_twentysix)
            ma_twentysix.append(ema_twentysix)
            ma_twelve.append(ema_twelve)

            if i >= 9:
                ema_signal = (macd[i] * (2/10)) + (signal[i - 1] * (1 - (2/10)))
            else:
                ema_signal = ema_twelve - ema_twentysix
            signal.append(ema_signal)
            


        #add analysis to dataframe
        df.insert(6, 'OBV', obv, True)
        df.insert(7, 'A/D', ad, True)
        df.insert(8, 'AROON UP', aroon_up, True)
        df.insert(9, 'AROON DOWN', aroon_down, True)
        df.insert(10, 'RSI', rsi, True)
        df.insert(11, 'MACD', macd, True)
        df.insert(12, 'SIGNAL', signal, True)

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
            sg.Button("Done"),
            
            sg.Listbox(values = [
            "Line",
            "Candle Stick"],
            key = "Style",
            size = (12,2))]
        ]
        graph_window = sg.Window("Stock Look up", layout)

        while True:
            graph_event, graph_values = graph_window.read()

            if graph_event == "Done":
                window.close()
                break

            #visualization
            fig, axes = plt.subplots(2, 1, figsize = (20, 15))
            if graph_event == "OBV":
                vres = sns.lineplot(ax = axes[0], data = df["OBV"], legend = 'auto')
            elif graph_event == "A/D":
                vres = sns.lineplot(ax = axes[0], data = df["A/D"], legend = 'auto')
            elif graph_event == "AROON":
                vres = sns.lineplot(ax = axes[0], data = df["AROON UP"], legend = 'full', label = "UP")
                vres = sns.lineplot(ax = axes[0], data = df["AROON DOWN"], legend = 'full', label = "DOWN")
            elif graph_event == "RSI":
                axes[0].axhline(y = 30, color = 'green', label = 'oversold', ls = '--') 
                axes[0].axhline(y = 70, color = 'red', label = 'overbought', ls = '--')
                vres = sns.lineplot(ax = axes[0], data = df["RSI"], legend = 'auto', label = "RSI")      
            elif graph_event == "MACD":
                axes[0].axhline(y = 0, color = 'red', ls = '--')
                vres = sns.lineplot(ax = axes[0], data = df["MACD"], legend = 'full', label = "MACD")
                vres = sns.lineplot(ax = axes[0], data = df["SIGNAL"], legend = 'full', label = "SIGNAL")

                
            if len(graph_values["Style"]) > 0:
                if graph_values["Style"][0] == "Line":
                    res1 = sns.lineplot(ax = axes[1], data = df["CLOSE"], legend = 'full')

                else:
                    candle = pd.DataFrame(data = date, columns = ["Date"])
                    candle.insert(1, 'Open', open, True)
                    candle.insert(2, 'High', high, True)
                    candle.insert(3, 'Low', low, True)
                    candle.insert(4, 'Close', close, True)
                    candle.set_index("Date", inplace = True)
                    
                    mpl.plot(candle, type = 'candle', ax = axes[1], style = 'yahoo')

            else:
                res1 = sns.lineplot(ax = axes[1], data = df["CLOSE"], legend = 'full')
                
            plt.show()

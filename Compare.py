import os
import tkinter as tk
import tkinter.ttk as ttk
import feedparser
import requests
import pandas as pd
from bs4 import BeautifulSoup
import feedparser
import numpy as np
from datetime import datetime
import re
import random
import string
from openpyxl import Workbook

def button_event():
     #print(mycombobox.get()) 
     rss_url = 'https://www.notebookcheck.net/RSS-Feed-Notebook-Reviews.8156.0.html'
 
     # 抓取資料
     rss = feedparser.parse(rss_url)
     r2 = requests.get(rss_url) 
     soup2 = BeautifulSoup(r2.text,"html.parser") 
     sel = soup2.select("pubDate")
        
     newtime=sel
     np.savetxt("result.txt", newtime, fmt = '%s')
     file=open('result.txt')
     lines=file.readlines()
     dt=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
     
     b = [i for i in lines if mycombobox.get() in i]
     #print(b)r"C:\Users\\"+
     save_index=[]
     #path = 'C:/Users/'+filename
     #assert os.path.isfile(path)
    
     if b :
            
        for j in range(0, len(b)):
             save_index=lines.index(b[j])
             #print(save_index)
             r = requests.get(rss.entries[lines.index(b[j])]['link'])
             soup= BeautifulSoup(r.text,"html.parser")  
             salt=''.join(random.sample(string.ascii_letters + string.digits, 8))
             filename=salt+'.xlsx'
             basic_INF_data=pd.DataFrame()
             R23M_data=pd.DataFrame()
             R23S_data=pd.DataFrame()
             R23MS_data_data=pd.DataFrame()
             R20M_data=pd.DataFrame()
             R20S_data=pd.DataFrame()
             R20MS_data_data=pd.DataFrame()
             Dmark_data=pd.DataFrame()
             
             with pd.ExcelWriter(filename, engine='openpyxl') as writer:
             
                 basic_INF= soup.find_all('table',{'class': 'contenttable'})
                 if basic_INF:
                     basic_INF_data=pd.read_html(str(basic_INF),header=0)
                     basic_INF_data=pd.concat(basic_INF_data, axis=0, ignore_index=False)
                     basic_INF_data.to_excel(writer, index=0, sheet_name='Competitors in Comparison')
                 else:
                     print("basic_INF_data not exsit")

                 R23M= soup.find_all(attrs={'class': re.compile("r_compare_bars r_compare_benchmark_768_2370")})
                 R23S= soup.find_all(attrs={'class': re.compile("r_compare_bars r_compare_benchmark_768_2371")})
                 R23MS= soup.find_all(attrs={'class': re.compile("r_compare_bars r_compare_benchmark_768")}) 
                
                 if R23M and R23S:
                     R23M_data=pd.read_html(str(R23M),header=0)
                     R23M_data=pd.concat(R23M_data, axis=0, ignore_index=False)
                     R23M_data.to_excel(writer, index=0, sheet_name='R23M')
                    
                     R23S_data=pd.read_html(str(R23S),header=0)
                     R23S_data=pd.concat(R23S_data, axis=0, ignore_index=False)
                     R23S_data.to_excel(writer, index=0, sheet_name='R23S')
                 elif R23MS :
                     R23MS_data=pd.read_html(str(R23MS),header=0)
                     R23MS_data=pd.concat(R23MS_data, axis=0, ignore_index=False)
                     R23MS_data.to_excel(writer, index=0, sheet_name='R23MS')
                 else:
                     print("R23MS_data not exsit")
                

                 R20M= soup.find_all(attrs={'class': re.compile("r_compare_bars r_compare_benchmark_671_2014")}) 
                 R20S= soup.find_all(attrs={'class': re.compile("r_compare_bars r_compare_benchmark_671_2015")})
                 R20MS= soup.find_all(attrs={'class': re.compile("r_compare_bars r_compare_benchmark_671")})
                 
                 if R20M and R20S:
                     R20M_data=pd.read_html(str(R20M),header=0)
                     R20M_data=pd.concat(R20M_data, axis=0, ignore_index=False)
                     R20M_data.to_excel(writer, index=0, sheet_name='R20M')
                        
                     R20S_data=pd.read_html(str(R20S),header=0)
                     R20S_data=pd.concat(R20S_data, axis=0, ignore_index=False)
                     #print(R20S_data)
                     R20S_data.to_excel(writer, index=0, sheet_name='R20S')
                 elif R20MS:
                     R20MS_data=pd.read_html(str(R20MS),header=0)
                     R20MS_data=pd.concat(R20MS_data, axis=0, ignore_index=False)
                     R20MS_data.to_excel(writer, index=0, sheet_name='R20MS')
                 else:
                     print("R23MS_data not exsit")
              

                 Dmark= soup.find_all(attrs={'class': re.compile("r_compare_bars r_compare_benchmark_201")})
                 if Dmark:
                     Dmark_data=pd.read_html(str(Dmark),header=0)
                     Dmark_data=pd.concat(Dmark_data, axis=0, ignore_index=False)
                     #print(Dmark_data)
                     Dmark_data.to_excel(writer, index=0, sheet_name='3Dmark')
                 else:
                     print("Dmark_data not exsit")
                    
                 if basic_INF or R23M or R23S or R23MS or R20M or R20S or R20MS or Dmark:
                    writer.save()
                    writer.close()
                    print(b[j]+"Cache successful")
                 else:
                    df = pd.DataFrame({"a": ["此網站資料不存在符合資料"]})
                    df.to_excel(writer, startrow=1, startcol=0)
                    writer.close()
                    print(b[j]+"There's no data") 
                 
     else:
         w= tk.Label(root, text="該月份資料不存在")
         w.pack()
         print("該月份資料不存在")
         

    
root = tk.Tk()
root.title('my window')
root.geometry('300x200')

comboboxList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
mycombobox = ttk.Combobox(root, state='readonly')
mycombobox['values'] = comboboxList

mycombobox.pack(pady=10)
mycombobox.current(0)


buttonText =  tk.StringVar()
buttonText.set('button')
tk.Button(root, textvariable=buttonText, command=button_event).pack()


root.mainloop()   
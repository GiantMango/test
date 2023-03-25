from time import sleep
from random import randint
import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import numpy as np

######  state  ######
grab = False
first_col = True
first = True
page = 1
data = np.empty(7)
final_p = 285
add_row = True
xy = 0
first_open = True
excel_mode = 'w'

###### 請求表頭  ######
headers = {
    'content-type': 'text/html; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36(KHTML, like Gecko) Chrome/76.0.3809.132 Safari/537.36'
           }


######  網址list  ######
addr_ls = []
for p in range(1, final_p):
    addr = 'https://civil.gov.taipei/News.aspx?n=21E85BBF867C4A19&sms=59AD6E6606F6002F&page=' + str(p) + '&PageSize=20'
    addr_ls.append(addr)

###### for address loop ######
for addr in addr_ls:
    res = requests.get(addr, headers=headers)
    soup = BeautifulSoup(res.text, "lxml")
    data_con = data.nbytes
    temp_data = []
    first_col = True


######  只抓column title  ######
    while grab is False:
        col = []
        for title in soup.find_all('td', limit=7):
            col.append(title.get('data-title'))
        grab = True
    
######  check memory usage  ######
    if data_con > 3e+9:
        print("Write to excel")  
        temp = pd.DataFrame(data)
        df = pd.DataFrame(temp.transpose())
        
        
        if first_open == False:
            excel_mode = 'a'

        with pd.ExcelWriter('台北市公民參與網.xlsx', mode=excel_mode) as writer:
            df.to_excel(writer, sheet_name='test1', index=False)

        first_open = False
        
        data = np.array([[]])
        print("Cleared memory")
        first = True
        
    else:
        print("memory usage", data_con, "bytes")
        

        
######  抓data  ######
        for i in range(7):

            temp_col = []
            
            for item in soup.find_all(class_='CCMS_jGridView_td_Class_'+str(i)):
                temp_col.append(item.string)
                
            if first_col == True:
                temp_data = np.array([temp_col])
            else:
                temp_data = np.append(temp_data, [temp_col], axis = xy)

            sleep(randint(1,2))
            first_col = False
            
        if first == True:
            data = temp_data
            first = False
        else:
            data = np.append(data, temp_data, axis = 1)
            
######  open csv  ######
    print("Page Done", page)
    page += 1


########  寫到excel ######
##temp = pd.DataFrame(data)
##df = pd.DataFrame(temp.transpose())
##df.columns = col
##print("Writing to excel...")
##df.to_excel('台北市公民參與網.xlsx', sheet_name='test1', index=False)

temp = pd.DataFrame(data)
df = pd.DataFrame(temp.transpose())
df.columns = col

if first_open == False:
    excel_mode = 'a'
    
with pd.ExcelWriter('台北市公民參與網.xlsx', mode=excel_mode) as writer:
    df.to_excel(writer, sheet_name='test1', index=False)

print("Crawling Done")

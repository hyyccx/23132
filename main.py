import requests
from bs4 import BeautifulSoup
import json
import pandas as pd
import openpyxl
import datetime
data=[]
now = datetime.datetime.now()
txt = '上次更新時間為：' + str(now) + '\n'

# 打開csv並且暫時叫csvlogFile
with open('log.csv', 'a', newline='') as csvlogFile:
    # 寫入資料
    csvlogFile.write(txt)

# 關閉讀取csv
csvlogFile.close

url = "https://www.cnyes.com/twstock/006208/summary/overview"
headers = {'User-Agent' : 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Mobile Safari/537.36'}
res = requests.get(url,headers=headers)
soup = BeautifulSoup(res.text,"html.parser")
#當天最後價格
articles = soup.find("h3",class_="jsx-162737614")
articles = articles.text
print(articles)
#當天日期
dataa = soup.find_all("span",class_="jsx-162737614")
dataa = dataa[4].text
print(dataa)
#年均價抓取
url2 = "https://www.cnyes.com/archive/twstock/technical/006208.htm"
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36'}
res = requests.get(url2, headers=headers)
soup = BeautifulSoup(res.text, "html.parser")
ave = soup.find_all("td",class_="rt")
ave = ave[7].text
print(ave)

date = {"當天日期":dataa,"當天最後價格": articles ,"年均價":ave}
data.append(date)
df = pd.DataFrame(data)
print(df)
#df.to_excel("stock_everyday.xlsx", index= False ,  engine = "openpyxl")

file_path = "stock_everyday.xlsx"
# 打开现有的 Excel 文件或创建新的
try:
    wb = openpyxl.load_workbook(file_path)
except FileNotFoundError:
    wb = openpyxl.Workbook()
# 选择工作表
s1 = wb.active
# 将整个 DataFrame 添加到工作表
for index, row in df.iterrows():
    s1.append(row.tolist())
# 保存更改
wb.save(file_path)



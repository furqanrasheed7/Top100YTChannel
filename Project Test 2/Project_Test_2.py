#Top 10 Youtube Channels through WebScrapping

import requests
from bs4 import BeautifulSoup
import numpy
import xlsxwriter

url = 'https://us.youtubers.me/global/all/top-300-youtube-channels'
source = requests.get(url)
source.raise_for_status()

soup = BeautifulSoup(source.text,'html.parser')

tab_ele = soup.find('table', class_='top-charts')

channels = tab_ele.find_all('tr')
i=0
details = [''] * 2150

for channel in tab_ele.find_all('td'):
    i = i+1
    check=channel.text
    details[i] = check

for i in range(2110):
    details[i] = details[i].strip()
    print(details[i])

workbook = xlsxwriter.Workbook("Project Test 2.xlsx")
workbook = xlsxwriter.Workbook('C:/Users/Furqan Rasheed/source/repos/Project Test 2 Solution/Project Test 2/Project Test 2 Excel.xlsx')
worksheet = workbook.add_worksheet("FirstSheet")

worksheet.write(0,0, 'Rank')
worksheet.write(0,1, 'Channel Name')
worksheet.write(0,2, 'Subscribers')
worksheet.write(0,3, 'Total Views')
worksheet.write(0,4, 'Total Videos')
worksheet.write(0,5, 'Category')
worksheet.write(0,6, 'Year Started')

for j in range(300):
    i=1+(j*7)
    worksheet.write(j+1,0, details[i])
    worksheet.write(j+1,1, details[i+1])
    worksheet.write(j+1,2, details[i+2])
    worksheet.write(j+1,3, details[i+3])
    worksheet.write(j+1,4, details[i+4])
    worksheet.write(j+1,5, details[i+5])
    worksheet.write(j+1,6, details[i+6])



workbook.close()

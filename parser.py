import xlwt, re, scrapy, requests
from datetime import datetime
from bs4 import BeautifulSoup


# xls and parsing settings
name = 'techno-gear'
curr_data = datetime.now().strftime("%d-%m-%Y %H:%M")
url = 'http://pp43.ru/'
center = xlwt.easyxf("align: horiz center")
# end xls and parsing settings 


# xls title init
wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
ws.write_merge(0, 0, 0, 4, name + ' ' + curr_data, center)
ws.write(1, 0, 'Название', center)
ws.write(1, 1, 'Производитель', center)
ws.write(1, 2, 'Наличие / Сроки поставки', center)
ws.write(1, 3, 'Цена, руб.', center)
ws.write(1, 4, '')
# end xls title init 


# parsing
headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:45.0) Gecko/20100101 Firefox/45.0'
      }
r = requests.get(url, headers = headers)
with open('test.html', 'w', encoding="utf-8") as output_file:
  output_file.write(r.text)
# end parsing 



wb.save(f'{name}.xls')
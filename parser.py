import xlwt, re, scrapy, requests, chardet
from datetime import datetime
from lxml import html

# xls and parsing settings
name = 'techno-gear-utop'
curr_data = datetime.now().strftime("%d-%m-%Y %H:%M")
url = 'https://pp43.ru/uplotneniya_gidrotsilindrov?page='
center = xlwt.easyxf("align: horiz center")
# end xls and parsing settings 
row = 1


# xls title init
wb = xlwt.Workbook()
ws = wb.add_sheet('products', cell_overwrite_ok=True)
ws.write_merge(0, 0, 0, 4, name + ' ' + curr_data, center)
ws.write(1, 0, 'Название', center)
ws.write(1, 1, 'Производитель', center)
ws.write(1, 2, 'Наличие / Сроки поставки', center)
ws.write(1, 3, 'Цена, руб.', center)
# end xls title init 


# parsing
def loadPage(url):
  headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.9; rv:45.0) Gecko/20100101 Firefox/45.0'}
  r = requests.get(url, headers = headers)
  return r.text


def pageParse(page):
  html_tree = html.fromstring(page)
  list = html_tree.xpath('//div[@class="product-list"]')
  if list: list = list[0].xpath('.//a[@class="quick_link"]/@href')
  return list
  
def productParse(list):
  global row
  domain = 'http://pp43.ru/'
  for link in list:

    html_tree = html.fromstring(loadPage(domain + link))
    name = html_tree.xpath('//h2[@class="product-name"]/a/text()')[0]
    manufacturer = ''
    if (len(html_tree.xpath('//div[@class="description"]/a/text()')) > 0):
      manufacturer = html_tree.xpath('//div[@class="description"]/a/text()')[0]
    availability = html_tree.xpath('//div[@class="description"]/span/text()')[-1]
    price = html_tree.xpath('//span[@class="price-new"]/span/text()')[0]
    re_price = re.search(r'[0-9\.]+', price).group(0)
    
    row += 1
    ws.write(row, 0, name)
    ws.write(row, 1, manufacturer)
    ws.write(row, 2, availability)
    ws.write(row, 3, re_price)
    


def parsingLoop():
  page = 1;

  while(True):
    currPage = loadPage(url + str(page))
    product_list = pageParse(currPage)
    if len(product_list) == 0: break

    productParse(product_list)

    page += 1

  

parsingLoop()
# end parsing 



wb.save(f'{name}.xls')
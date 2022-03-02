import openpyxl
#import urllib.request
from fake_useragent import UserAgent
import requests
from lxml import html
from requests_html import HTMLSession

#file1 = open("MyFile.txt","w")
path = input('Please enter path of excel file')

if (path == ''):
    path = "INVENTORY.xlsx"

print('Loading the file... Please wait')
#wb_obj = openpyxl.load_workbook(path)
#print(wb_obj.sheetnames)

numbers = ["270336923202","270336993632","270337237905","270336057802","270337235902","270336921840","270336923875","270336056872","270336056541","270336991467","270336056460","270336923875","270336918236","270336056460"]

print(numbers)

url = ''

for number in numbers:
    url = url+number+','

url = url.rstrip(url[-1])

url = 'https://www.fedex.com/fedextrack/summary?trknbr='+url
print(url)

url = 'https://www.fedex.com/fedextrack/summary?trknbr=270336923202'

#ua = UserAgent()
#print(ua.chrome)
#header = {'User-Agent':str(ua.chrome)}
#print(header)
#url = "https://www.hybrid-analysis.com/recent-submissions?filter=file&sort=^timestamp"
#htmlContent = requests.get(url, headers=header)
#tree = html.fromstring(htmlContent.text)

#prices = tree.xpath('//div[@class="shipment-status__key"]')

#print(html.tostring(prices))

#s = HTMLSession()
#response = s.get(url)
#response.html.render(wait=10, sleep=10)

#print(response.text)

#print(tree.content)

import requests
import openpyxl
from bs4 import BeautifulSoup
page = requests.get("https://www.mirea.ru/schedule/")
soup = BeautifulSoup(page.text, "html.parser")
result = soup.find("div", {"class":"rasspisanie"}).\
    find(string = "Институт информационных технологий").\
    find_parent("div").\
    find_parent("div").\
    find_all('a', class_='uk-link-toggle')

links = []

for x in result:
    links.append(x['href'])

i = 1
"""
for link in links:
    if "/ИИТ_1" in link: # среди всех ссылок найти нужную
        print (link)
        f = open('kurs_1.xlsx', "wb") # открываем файл для записи, в режиме wb
        resp = requests.get(link) # запрос по ссылке
        f.write(resp.content)
        i+=1
"""
book = openpyxl.load_workbook("kurs_1.xlsx")  # открытие файла
sheet = book.active  # активный лист
print(sheet['BN2'].value)
import notMain
import requests
import xml.etree.ElementTree as ET
import pandas as pd
"""blankadadadadaadadada"""
file_name = "vacancies_dif_currencies.csv"
data = notMain.DataSet(file_name, "none")
data = data.csv_reader()
currency_dic_count = {}

for item in data[1]:
    if item[3] == "":
        continue
    if item[3] not in currency_dic_count:
        currency_dic_count[item[3]] = 1
    else:
        currency_dic_count[item[3]] += 1

big_currency = []
for item in currency_dic_count:
    if currency_dic_count[item] >= 5000:
        big_currency.append(item)

data_dict = {}
for item in big_currency:
    if item != "RUR":
        data_dict[item] = []

index = []
for year in range(2005, 2023):
    for j in range(1, 13):
        if j < 10:
            month = '0' + str(j)
        else:
            month = j
        url = f'http://www.cbr.ru/scripts/XML_daily.asp?date_req=01/{month}/{year}'
        r = requests.get(url)
        with open('XML_daily.xml', 'wb') as file:
            file.write(r.content)
        tree = ET.parse('XML_daily.xml')
        root = tree.getroot()
        index.append(f'{year}-{month}')
        for item in root:
            currency = item.find('CharCode').text
            money = item.find('Value').text.replace(',', '.')
            money = float(money) / float(item.find('Nominal').text)
            if currency in data_dict:
                data_dict[currency].append(money)
        for item in data_dict:
            if len(data_dict[item]) != len(index):
                data_dict[item].append('NONE')

df = pd.DataFrame(data=data_dict, index=index)
df.to_csv('currency.csv')
print(df)
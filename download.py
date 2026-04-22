import csv
import requests
from collections import defaultdict

# Твои данные авторизации
auth = ('твой_логин', 'твой_пароль')

# 1. Скачиваем остатки
r_stock = requests.get(
    'http://193.32.203.98:8153/api.rsc/current_stock_full?@CSV&$select=product_no,SPB_STOCK&$filter=SPB_STOCK gt 0',
    auth=auth
)
stock_data = list(csv.DictReader(r_stock.text.splitlines()))

# 2. Скачиваем цены
r_price = requests.get(
    'http://193.32.203.98:8153/api.rsc/116179_pl?@CSV&$select=product_no,price',
    auth=auth
)
price_data = {row['product_no']: row['price'] for row in csv.DictReader(r_price.text.splitlines())}

# 3. Соединяем
result = []
for item in stock_data:
    result.append({
        'product_no': item['product_no'],
        'SPB_STOCK': item['SPB_STOCK'],
        'price': price_data.get(item['product_no'], 'нет цены')
    })

# 4. Сохраняем результат
with open('result.csv', 'w', newline='') as f:
    writer = csv.DictWriter(f, fieldnames=['product_no', 'SPB_STOCK', 'price'])
    writer.writeheader()
    writer.writerows(result)

print('Готово! Результат в result.csv')
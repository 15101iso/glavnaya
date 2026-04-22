import csv
import requests
from collections import defaultdict
import os

# ===== НАСТРОЙКИ =====
LOGIN = '116179'      # <--- ВВЕДИ СВОЙ ЛОГИН
PASSWORD = 'wvKnilTAJEe6G4xD44Ji'  # <--- ВВЕДИ СВОЙ ПАРОЛЬ
# =====================

# Базовый URL API
BASE_URL = 'http://193.32.203.98:8153/api.rsc'

# Авторизация
auth = (LOGIN, PASSWORD)

print('🚀 Начинаем выгрузку...')

try:
    # 1. 📦 СКАЧИВАЕМ ОСТАТКИ (только СПб > 0)
    print('📦 Скачиваем остатки СПб...')
    stock_url = f'{BASE_URL}/current_stock_full?@CSV&$select=product_no,SPB_STOCK&$filter=SPB_STOCK gt 0'
    r_stock = requests.get(stock_url, auth=auth)
    r_stock.raise_for_status()
    
    # Парсим CSV
    stock_lines = r_stock.text.strip().splitlines()
    stock_reader = csv.DictReader(stock_lines)
    stock_data = list(stock_reader)
    print(f'   ✅ Найдено {len(stock_data)} товаров на складе СПб')

    # 2. 💰 СКАЧИВАЕМ ЦЕНЫ
    print('💰 Скачиваем цены...')
    price_url = f'{BASE_URL}/116179_pl?@CSV&$select=product_no,price'
    r_price = requests.get(price_url, auth=auth)
    r_price.raise_for_status()
    
    # Парсим CSV и делаем словарь {артикул: цена}
    price_lines = r_price.text.strip().splitlines()
    price_reader = csv.DictReader(price_lines)
    price_dict = {row['product_no']: row['price'] for row in price_reader}
    print(f'   ✅ Найдено {len(price_dict)} цен')

    # 3. 🔗 СОЕДИНЯЕМ ДАННЫЕ
    print('🔗 Соединяем данные...')
    result = []
    not_found = 0
    
    for item in stock_data:
        product_no = item['product_no']
        price = price_dict.get(product_no)
        
        if price is None:
            not_found += 1
            price = 'Цена не найдена'
            
        result.append({
            'product_no': product_no,
            'SPB_STOCK': item['SPB_STOCK'],
            'price': price
        })

    # 4. 💾 СОХРАНЯЕМ РЕЗУЛЬТАТ
    output_file = 'ostanki_spb_s_cenami.csv'
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['product_no', 'SPB_STOCK', 'price'])
        writer.writeheader()
        writer.writerows(result)
    
    print(f'\n✅ ГОТОВО! Файл сохранён: {output_file}')
    print(f'   📊 Всего записей: {len(result)}')
    if not_found > 0:
        print(f'   ⚠️ Товаров без цены: {not_found}')
    
    # Покажем первые 5 строк
    print('\n📋 Первые 5 строк:')
    for i, row in enumerate(result[:5]):
        print(f'   {row["product_no"]} | СПб: {row["SPB_STOCK"]} | Цена: {row["price"]}')

except requests.exceptions.RequestException as e:
    print(f'\n❌ Ошибка при запросе к API: {e}')
except Exception as e:
    print(f'\n❌ Неожиданная ошибка: {e}')
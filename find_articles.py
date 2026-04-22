import csv
import os
from datetime import datetime

def find_articles_from_list(articles_file, database_file, output_file=None):
    """
    Ищет артикулы из первого файла во втором файле и сохраняет результаты
    
    Args:
        articles_file (str): файл со списком артикулов для поиска (xr250.csv)
        database_file (str): файл с базой данных (allmrt.csv)
        output_file (str): выходной файл с результатами
    """
    
    # Проверяем существование файлов
    if not os.path.exists(articles_file):
        print(f"❌ Ошибка: Файл '{articles_file}' не найден!")
        return False
    
    if not os.path.exists(database_file):
        print(f"❌ Ошибка: Файл '{database_file}' не найден!")
        return False
    
    # Создаем имя выходного файла, если не указано
    if output_file is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_name = os.path.splitext(articles_file)[0]
        output_file = f'{base_name}_found_{timestamp}.csv'
    
    print("=" * 80)
    print(f"🔍 ПОИСК АРТИКУЛОВ ИЗ ФАЙЛА: {articles_file}")
    print(f"📁 В БАЗЕ ДАННЫХ: {database_file}")
    print("=" * 80)
    
    # ШАГ 1: Читаем список артикулов для поиска
    articles_to_find = set()
    articles_file_delimiter = None
    
    # Пробуем разные разделители для файла с артикулами
    delimiters = [',', ';', '\t', '|']
    
    for delimiter in delimiters:
        try:
            with open(articles_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=delimiter)
                
                # Проверяем первую строку
                first_row = next(reader)
                if first_row and len(first_row) > 0:
                    # Предполагаем, что артикулы в первой колонке
                    for row in [first_row] + list(reader):
                        if row and len(row) > 0 and row[0].strip():
                            articles_to_find.add(str(row[0]).strip())
                    
                    articles_file_delimiter = delimiter
                    print(f"✅ Файл с артикулами прочитан (разделитель: '{delimiter}')")
                    print(f"📊 Найдено артикулов для поиска: {len(articles_to_find)}")
                    break
        except:
            continue
    
    if not articles_to_find:
        print("❌ Не удалось прочитать артикулы из файла!")
        return False
    
    # ШАГ 2: Читаем базу данных и ищем совпадения
    database_headers = []
    found_results = []
    database_delimiter = None
    article_col_index = None
    
    print(f"\n🔎 Поиск в базе данных...")
    
    for delimiter in delimiters:
        try:
            with open(database_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=delimiter)
                
                # Читаем заголовки
                database_headers = next(reader)
                
                # Ищем столбец с артикулами (Article)
                for i, header in enumerate(database_headers):
                    header_lower = header.strip().lower()
                    if 'article' in header_lower or 'артикул' in header_lower:
                        article_col_index = i
                        print(f"✅ Найден столбец с артикулами: '{header}' (индекс {i})")
                        break
                
                if article_col_index is None:
                    continue
                
                database_delimiter = delimiter
                
                # Поиск совпадений
                for row_num, row in enumerate(reader, 2):  # начинаем с 2 (после заголовков)
                    if len(row) > article_col_index:
                        article_value = str(row[article_col_index]).strip()
                        
                        # Проверяем, есть ли этот артикул в нашем списке
                        if article_value in articles_to_find:
                            found_results.append({
                                'row_num': row_num,
                                'article': article_value,
                                'data': row
                            })
                
                print(f"✅ База данных прочитана (разделитель: '{delimiter}')")
                print(f"📊 Всего строк в базе: {row_num - 1}")
                break
                
        except Exception as e:
            continue
    
    if article_col_index is None:
        print("❌ Не удалось найти столбец с артикулами в базе данных!")
        return False
    
    print(f"\n🎯 Найдено совпадений: {len(found_results)}")
    
    if not found_results:
        print("\n⚠️ Совпадений не найдено!")
        return True
    
    # ШАГ 3: Сохраняем результаты
    with open(output_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter=database_delimiter)
        
        # Заголовки (добавляем номер строки из исходного файла)
        new_headers = ['Source_Row'] + database_headers + ['Searched_Article']
        writer.writerow(new_headers)
        
        # Данные
        for result in found_results:
            writer.writerow([result['row_num']] + result['data'] + [result['article']])
    
    print(f"\n✅ Результаты сохранены в файл: {output_file}")
    
    # Создаем краткий отчет
    summary_file = output_file.replace('.csv', '_summary.csv')
    with open(summary_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Article', 'Found_in_Rows'])
        
        # Группируем по артикулам
        from collections import defaultdict
        article_groups = defaultdict(list)
        for result in found_results:
            article_groups[result['article']].append(result['row_num'])
        
        for article, rows in article_groups.items():
            writer.writerow([article, len(rows), str(rows)])
    
    print(f"✅ Краткий отчет: {summary_file}")
    
    # Статистика
    print("\n" + "=" * 80)
    print("📊 СТАТИСТИКА ПОИСКА")
    print("=" * 80)
    print(f"📌 Всего артикулов для поиска: {len(articles_to_find)}")
    print(f"📌 Найдено уникальных артикулов: {len(article_groups)}")
    print(f"📌 Всего найдено строк: {len(found_results)}")
    
    # Покажем первые несколько результатов
    if found_results:
        print("\n🔍 ПЕРВЫЕ 5 НАЙДЕННЫХ АРТИКУЛОВ:")
        for i, result in enumerate(found_results[:5], 1):
            print(f"  {i}. Строка {result['row_num']}: {result['article']}")
    
    return True

def find_articles_simple(articles_file, database_file, output_file='found_articles.csv'):
    """
    Упрощенная версия функции
    """
    
    # Читаем артикулы для поиска
    search_articles = set()
    with open(articles_file, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if row and row[0].strip():
                search_articles.add(row[0].strip())
    
    print(f"Найдено артикулов для поиска: {len(search_articles)}")
    
    # Ищем в базе данных
    results = []
    headers = []
    
    with open(database_file, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        headers = next(reader)
        
        # Находим столбец Article
        article_idx = None
        for i, h in enumerate(headers):
            if 'article' in h.lower():
                article_idx = i
                break
        
        if article_idx is None:
            print("Столбец Article не найден!")
            return
        
        # Поиск
        for row_num, row in enumerate(reader, 2):
            if len(row) > article_idx:
                article = row[article_idx].strip()
                if article in search_articles:
                    results.append((row_num, row, article))
    
    print(f"Найдено совпадений: {len(results)}")
    
    # Сохраняем
    with open(output_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Row_in_source'] + headers + ['Searched_article'])
        for row_num, row, article in results:
            writer.writerow([row_num] + row + [article])
    
    print(f"Результаты сохранены в {output_file}")

def main():
    """Основная функция"""
    
    # Настройки
    articles_file = 'xr250.csv'      # файл со списком артикулов
    database_file = 'allmrt.csv'      # файл с базой данных
    
    print("🔧 Параметры поиска:")
    print(f"   📁 Файл с артикулами: {articles_file}")
    print(f"   📁 База данных: {database_file}")
    print()
    
    # Запускаем поиск
    find_articles_from_list(articles_file, database_file)
    
    print("\n" + "=" * 80)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
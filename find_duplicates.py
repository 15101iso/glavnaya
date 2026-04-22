import csv
from collections import defaultdict
import os
from datetime import datetime

def find_duplicates_in_article(input_file, output_file=None):
    """
    Находит все повторяющиеся значения в столбце 'Article'
    и сохраняет результаты в CSV файл с ID строк
    
    Args:
        input_file (str): путь к входному CSV файлу
        output_file (str): путь к выходному CSV файлу (если не указан, создается автоматически)
    """
    
    # Проверяем существование файла
    if not os.path.exists(input_file):
        print(f"❌ Ошибка: Файл '{input_file}' не найден!")
        return False
    
    # Создаем имя выходного файла, если не указано
    if output_file is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'duplicates_article_{timestamp}.csv'
    
    print("=" * 80)
    print(f"🔍 ПОИСК ПОВТОРЕНИЙ В СТОЛБЦЕ 'ARTICLE'")
    print(f"📁 Входной файл: {input_file}")
    print("=" * 80)
    
    # Пробуем разные разделители
    delimiters = [',', ';', '\t', '|']
    used_delimiter = None
    article_col_index = None
    id_col_index = None
    headers = []
    all_rows = []
    
    for delimiter in delimiters:
        try:
            with open(input_file, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter=delimiter)
                
                # Читаем заголовки
                headers = next(reader)
                
                # Ищем столбцы Article и ID
                for i, header in enumerate(headers):
                    header_lower = header.strip().lower()
                    if 'article' in header_lower:
                        article_col_index = i
                        print(f"✅ Найден столбец Article: '{header}' (индекс {i})")
                    elif 'id' in header_lower or 'код' in header_lower or 'артикул' in header_lower:
                        id_col_index = i
                        print(f"✅ Найден столбец ID: '{header}' (индекс {i})")
                
                if article_col_index is None:
                    print(f"❌ Столбец 'Article' не найден!")
                    continue
                
                used_delimiter = delimiter
                
                # Читаем все строки
                for row_num, row in enumerate(reader, 2):  # начинаем с 2, т.к. 1 - заголовки
                    if len(row) > max(article_col_index, id_col_index if id_col_index else 0):
                        all_rows.append((row_num, row))
                
                break  # успешно прочитали
                
        except Exception as e:
            continue
    
    if article_col_index is None:
        print("❌ Не удалось найти столбец 'Article' ни с одним разделителем!")
        return False
    
    print(f"\n📊 Всего строк с данными: {len(all_rows)}")
    
    # Группируем строки по значениям Article
    article_groups = defaultdict(list)
    
    for row_num, row in all_rows:
        article_value = row[article_col_index].strip() if row[article_col_index] else ''
        if article_value:  # пропускаем пустые значения
            article_groups[article_value].append((row_num, row))
    
    print(f"📊 Уникальных значений Article: {len(article_groups)}")
    
    # Находим повторения (значения, которые встречаются больше 1 раза)
    duplicates = {article: rows for article, rows in article_groups.items() if len(rows) > 1}
    
    print(f"📊 Значений с повторениями: {len(duplicates)}")
    
    if not duplicates:
        print("\n✅ ПОВТОРЕНИЙ НЕ НАЙДЕНО!")
        return True
    
    # Создаем выходной CSV файл
    with open(output_file, 'w', encoding='utf-8', newline='') as f:
        # Добавляем новые колонки к заголовкам
        new_headers = headers + ['Article_Value', 'Row_Number', 'Duplicate_Group_ID']
        writer = csv.writer(f, delimiter=used_delimiter)
        writer.writerow(new_headers)
        
        # Записываем все строки с повторениями
        group_id = 1
        rows_written = 0
        
        for article, rows in duplicates.items():
            for row_num, row in rows:
                writer.writerow(row + [article, row_num, group_id])
                rows_written += 1
            group_id += 1
    
    print(f"\n✅ Найдено групп повторений: {len(duplicates)}")
    print(f"✅ Всего строк с повторениями: {rows_written}")
    print(f"✅ Результат сохранен в: {output_file}")
    
    # Создаем дополнительный файл со статистикой
    stats_file = output_file.replace('.csv', '_stats.csv')
    with open(stats_file, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Article_Value', 'Count', 'Row_Numbers', 'Group_ID'])
        
        group_id = 1
        for article, rows in duplicates.items():
            row_numbers = [str(r[0]) for r in rows]
            writer.writerow([article, len(rows), ', '.join(row_numbers), group_id])
            group_id += 1
    
    print(f"✅ Статистика сохранена в: {stats_file}")
    
    return True

def main():
    """Основная функция"""
    
    input_file = 'allmrt.csv'
    
    # Запускаем поиск
    find_duplicates_in_article(input_file)
    
    print("\n" + "=" * 80)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
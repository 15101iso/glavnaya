def save_to_csv(original_file, column_name, duplicates, headers, all_rows, col_index):
    """
    Сохраняет результаты в CSV файлы (нет ограничений на колонки)
    """
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    
    try:
        print("\n📊 Создание CSV файлов...")
        
        # Собираем все строки с повторениями
        rows_with_duplicates = []
        all_duplicate_values = set(duplicates.keys())
        
        for row_num, row in enumerate(all_rows, 1):
            if row_num == 1:
                continue
                
            if len(row) > col_index:
                value = row[col_index].strip() if row[col_index] else ''
                if value in all_duplicate_values:
                    repeat_count = len(duplicates[value])
                    rows_with_duplicates.append((row_num, row, value, repeat_count))
        
        if rows_with_duplicates:
            # CSV файл со всеми строками
            csv_file1 = f'duplicates_{column_name}_all_rows_{timestamp}.csv'
            with open(csv_file1, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                # Заголовки
                writer.writerow(headers + ['Количество повторений'])
                # Данные
                for row_num, row, value, repeat_count in rows_with_duplicates:
                    writer.writerow(row + [repeat_count])
            print(f"✅ CSV файл (все строки): {csv_file1}")
            
            # CSV файл только с повторяющимися значениями
            csv_file2 = f'duplicates_{column_name}_values_only_{timestamp}.csv'
            with open(csv_file2, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Значение', 'Количество повторений', 'Номера строк'])
                
                sorted_duplicates = sorted(duplicates.items(), key=lambda x: len(x[1]), reverse=True)
                for value, rows in sorted_duplicates:
                    writer.writerow([value, len(rows), str(rows)])
            
            print(f"✅ CSV файл (только значения): {csv_file2}")
            
    except Exception as e:
        print(f"⚠️ Ошибка при создании CSV файлов: {e}")
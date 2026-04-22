import csv

# Имя файла
filename = 'allMRT.csv'

try:
    with open(filename, 'r', encoding='utf-8') as file:
        reader = csv.reader(file)
        
        # Читаем первую строку - заголовки
        headers = next(reader)
        
        print("=" * 60)
        print(f"Файл: {filename}")
        print("=" * 60)
        print(f"Названия столбцов ({len(headers)}):")
        print("-" * 60)
        
        # Выводим каждый заголовок с номером
        for i, header in enumerate(headers, 1):
            print(f"{i:3}. {header}")
            
        print("-" * 60)
        
        # Покажем пример первых 3 строк данных
        print("\nПример первых 3 строк данных:")
        print("-" * 60)
        
        row_count = 0
        for row in reader:
            if row_count < 3:
                # Ограничим вывод для читаемости
                if len(row) > 5:
                    preview = row[:5] + ['...'] + row[-2:]
                else:
                    preview = row
                print(f"Строка {row_count + 2}: {preview}")
                row_count += 1
            else:
                break
                
        print("-" * 60)
        
except FileNotFoundError:
    print(f"Ошибка: Файл '{filename}' не найден!")
    print("Убедитесь, что файл находится в той же папке, что и программа.")
except Exception as e:
    print(f"Произошла ошибка: {e}")
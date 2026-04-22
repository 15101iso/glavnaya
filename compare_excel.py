"""
Программа для удаления повторяющихся значений в столбце
Оставляет только значения, которые встречаются 1 раз
"""

import pandas as pd
import os

def keep_unique_only(input_file, column_name=None, output_file=None):
    """
    Оставляет только значения, которые встречаются 1 раз
    
    Parameters:
    input_file (str): путь к файлу
    column_name (str): название столбца (если None - используется первый столбец)
    output_file (str): имя выходного файла
    """
    
    try:
        # Проверяем существование файла
        if not os.path.exists(input_file):
            print(f"Ошибка: файл {input_file} не найден!")
            return False
        
        # Читаем файл
        print(f"Чтение файла {input_file}...")
        df = pd.read_excel(input_file)
        
        print(f"\nИсходная информация:")
        print(f"Столбцы в файле: {list(df.columns)}")
        print(f"Всего строк: {len(df)}")
        
        # Если имя столбца не указано, берем первый
        if column_name is None:
            column_name = df.columns[0]
            print(f"\nИспользуется столбец: {column_name}")
        
        # Проверяем, существует ли столбец
        if column_name not in df.columns:
            print(f"Ошибка: столбец '{column_name}' не найден!")
            print(f"Доступные столбцы: {list(df.columns)}")
            return False
        
        # Анализируем частоту значений
        print(f"\nАнализ столбца '{column_name}':")
        
        # Считаем частоту каждого значения
        value_counts = df[column_name].value_counts()
        
        # Находим значения, которые встречаются только 1 раз
        unique_values = value_counts[value_counts == 1].index.tolist()
        
        # Находим значения, которые встречаются несколько раз
        duplicate_values = value_counts[value_counts > 1].index.tolist()
        
        # Статистика
        total_values = len(df[column_name].dropna())
        unique_count = len(unique_values)
        duplicate_count = len(duplicate_values)
        
        print(f"\nСтатистика:")
        print(f"Всего значений в столбце: {total_values}")
        print(f"Уникальных значений (встречаются 1 раз): {unique_count}")
        print(f"Значений с повторениями (2+ раз): {duplicate_count}")
        
        if duplicate_count > 0:
            print(f"\nПримеры повторяющихся значений (первые 10):")
            for i, val in enumerate(duplicate_values[:10]):
                count = value_counts[val]
                print(f"  {val}: встречается {count} раз")
        
        # Фильтруем: оставляем только те строки, где значение встречается 1 раз
        df_filtered = df[df[column_name].isin(unique_values)].copy()
        
        print(f"\nРезультат:")
        print(f"Осталось строк: {len(df_filtered)}")
        print(f"Удалено строк: {len(df) - len(df_filtered)}")
        
        # Сохраняем результат
        if output_file is None:
            base_name = os.path.splitext(input_file)[0]
            output_file = f"{base_name}_unique_only.xlsx"
        
        df_filtered.to_excel(output_file, index=False)
        print(f"\n✓ Результат сохранен в: {output_file}")
        
        # Показываем первые строки результата
        if len(df_filtered) > 0:
            print(f"\nПервые 10 уникальных значений:")
            print(df_filtered[column_name].head(10).tolist())
        else:
            print("\nВнимание: не найдено уникальных значений!")
        
        return True
        
    except Exception as e:
        print(f"\n✗ ОШИБКА: {e}")
        return False

def main():
    """Главная функция"""
    print("="*60)
    print("ПРОГРАММА ДЛЯ УДАЛЕНИЯ ПОВТОРЯЮЩИХСЯ ЗНАЧЕНИЙ")
    print("Оставляем только те, которые встречаются 1 раз")
    print("="*60)
    print()
    
    input_file = "frog.xlsx"
    
    # Проверяем наличие файла
    if not os.path.exists(input_file):
        print(f"Ошибка: файл {input_file} не найден!")
        print(f"Пожалуйста, поместите файл {input_file} в текущую папку")
        print(f"Текущая папка: {os.getcwd()}")
        input("\nНажмите Enter для выхода...")
        return
    
    # Читаем файл для просмотра структуры
    try:
        df = pd.read_excel(input_file, nrows=5)
        print(f"Найден файл: {input_file}")
        print(f"Столбцы в файле: {list(df.columns)}")
        
        # Если несколько столбцов, спрашиваем какой обрабатывать
        if len(df.columns) > 1:
            print(f"\nВ файле несколько столбцов.")
            print("1 - Обработать первый столбец")
            print("2 - Выбрать столбец вручную")
            print("3 - Обработать все столбцы (создать отдельные файлы)")
            
            choice = input("Ваш выбор (1, 2 или 3): ").strip()
            
            if choice == "2":
                col_name = input(f"Введите название столбца из {list(df.columns)}: ").strip()
                keep_unique_only(input_file, column_name=col_name)
            elif choice == "3":
                # Обрабатываем каждый столбец отдельно
                for col in df.columns:
                    print(f"\n{'='*50}")
                    print(f"Обработка столбца: {col}")
                    output_file = f"frog_{col}_unique_only.xlsx"
                    keep_unique_only(input_file, column_name=col, output_file=output_file)
            else:
                keep_unique_only(input_file)
        else:
            keep_unique_only(input_file)
        
    except Exception as e:
        print(f"\n✗ ОШИБКА при чтении файла: {e}")
    
    print("\n" + "="*60)
    input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
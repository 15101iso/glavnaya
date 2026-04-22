mport pandas as pd

# Замените 'your_file.xls' на путь к вашему файлу
file_path = 'your_file.xls'

# Чтение файла
data = pd.read_excel(file_path)

# Вывод данных
print(data)

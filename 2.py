import pandas as pd

# Путь к файлу Excel
file_path = "output/Ник04.11-10.11.xlsx"

# Загрузка данных из Excel
df = pd.read_excel(file_path)

# Проверка типов данных, чтобы определить числовые столбцы
numeric_cols = df.select_dtypes(include=["number"]).columns

# Создание сводной таблицы с суммами по числовым полям
summary = df[numeric_cols].sum().reset_index()
summary.columns = ["Поле", "Сумма"]

# Сохранение результата в Excel
output_path = "summary_table.xlsx"
summary.to_excel(output_path, index=False)

print(f"Сводная таблица сохранена в {output_path}")


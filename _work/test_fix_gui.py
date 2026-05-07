#!/usr/bin/env python3
"""
Тест исправления метода _create_table_from_df.
Проверяем, что метод корректно обрабатывает массивы и не вызывает ValueError.
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import numpy as np

# Импортируем класс, но без Qt
# Создадим mock-класс, чтобы протестировать логику
from pbi_modules.gui import MainWindow

# Создадим тестовый DataFrame с разными типами значений
df = pd.DataFrame({
    'col1': [1, 2, None, 4],
    'col2': ['a', 'b', 'c', 'd'],
    'col3': [np.array([1,2]), [3,4,5], pd.Series([6,7]), None],
    'col4': [pd.NA, pd.NaT, np.nan, 42],
})

print("Тестовый DataFrame:")
print(df)
print("\nТипы значений в col3:")
for val in df['col3']:
    print(f"  {type(val)}: {val}")

# Создадим экземпляр MainWindow без GUI (не инициализируем Qt)
# Для этого нужно модифицировать __init__, но проще протестировать только метод
# Воспользуемся тем, что метод не зависит от self, кроме self.log
# Создадим mock-объект
class MockMainWindow:
    def log(self, msg):
        print(f"LOG: {msg}")

mock = MockMainWindow()
mock._create_table_from_df = MainWindow._create_table_from_df.__get__(mock, MainWindow)

print("\nВызов _create_table_from_df...")
try:
    table = mock._create_table_from_df(df, df_key='test')
    print("Успешно! Таблица создана.")
    # Проверим, что нет ошибок при рендеринге
    print("Количество строк в таблице:", table.rowCount())
    print("Количество столбцов в таблице:", table.columnCount())
except Exception as e:
    print(f"Ошибка: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\nТест пройден.")
#!/usr/bin/env python3
"""
Тест исправления метода _create_table_from_df.
Проверяем, что pd.isna не вызывает ValueError для массивов.
"""
import sys
import os
import pandas as pd
import numpy as np

# Имитируем логику из исправленного метода
def safe_isna(value):
    """Безопасная проверка на NA, как в исправленном коде."""
    is_na = False
    try:
        na_result = pd.isna(value)
        if hasattr(na_result, '__len__'):
            # pd.isna вернул массив (например, для Series/ndarray)
            # считаем, что значение не является NA в смысле скаляра
            is_na = False
        else:
            is_na = bool(na_result)
    except (ValueError, TypeError):
        # если pd.isna не может обработать тип, считаем не NA
        is_na = False
    return is_na

def determine_text(value):
    """Определение текста для ячейки, как в исправленном коде."""
    is_na = safe_isna(value)
    if is_na:
        return ""
    elif hasattr(value, '__len__') and not isinstance(value, str):
        # Массивоподобный объект (list, tuple, Series, ndarray)
        text = str(value)
        if len(text) > 200:
            text = text[:200] + '...'
        return text
    elif isinstance(value, str) and '\n' in value:
        text = value[:200] + ('...' if len(value) > 200 else '')
        return text
    else:
        return str(value)

# Тестовые значения
test_values = [
    42,
    None,
    np.nan,
    pd.NA,
    "обычная строка",
    "строка с\nпереносом",
    [1, 2, 3],
    np.array([4, 5, 6]),
    pd.Series([7, 8, 9]),
    [],
    {'key': 'value'},
]

print("Тестирование безопасной проверки NA и определения текста:")
for val in test_values:
    try:
        is_na = safe_isna(val)
        text = determine_text(val)
        print(f"{type(val).__name__:20} {repr(val)[:50]:50} is_na={is_na:5} text={text[:30]}")
    except Exception as e:
        print(f"{type(val).__name__:20} {repr(val)[:50]:50} ERROR: {e}")

# Тест с DataFrame, который вызывает ошибку
print("\n--- Тест с DataFrame ---")
df = pd.DataFrame({
    'scalar': [1, None, np.nan],
    'list': [[1,2], [3,4,5], []],
    'array': [np.array([6,7]), np.array([]), None],
    'series': [pd.Series([8,9]), pd.Series([]), pd.Series([10])],
})

print("DataFrame:")
print(df)

# Проверим каждую ячейку
errors = []
for row_idx in range(len(df)):
    for col_idx in range(len(df.columns)):
        value = df.iloc[row_idx, col_idx]
        try:
            text = determine_text(value)
        except Exception as e:
            errors.append((row_idx, col_idx, value, e))

if errors:
    print("\nОбнаружены ошибки:")
    for err in errors:
        print(f"  ({err[0]},{err[1]}) {type(err[2]).__name__}: {err[3]}")
else:
    print("\nВсе ячейки обработаны успешно.")

sys.exit(1 if errors else 0)
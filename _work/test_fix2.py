#!/usr/bin/env python3
"""Тест исправления ошибки AttributeError в _apply_lineage_focus."""
import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication

sys.path.insert(0, '.')

from pbi_modules.gui import PBITAnalyzerMainWindow

def test_filter_with_string_value():
    """Тест, что метод не падает при наличии строковых значений в словаре."""
    app = QApplication(sys.argv)
    window = PBITAnalyzerMainWindow()
    
    # Создаем тестовые данные, имитирующие результат анализа
    # Включаем строковое значение 'excel_path' (не DataFrame)
    test_dfs = {
        'measures': pd.DataFrame({
            '№ меры': ['M1', 'M2', 'M3'],
            'Мера': ['Мера1', 'Мера2', 'Мера3'],
            '№ таблицы': ['s1', 's2', 's3']
        }),
        'tables_ref': pd.DataFrame({
            '№ таблицы': ['s1', 's2', 's3'],
            'Имя таблицы': ['Таблица1', 'Таблица2', 'Таблица3']
        }),
        'lineage': pd.DataFrame({
            'Тип источника': ['Мера', 'Мера', 'Таблица'],
            'ID источника': ['M1', 'M2', 's1'],
            'Имя источника': ['Мера1', 'Мера2', 'Таблица1'],
            'Тип приемника': ['Визуал', 'Визуал', 'Столбец'],
            'ID приемника': ['VIS:1', 'VIS:2', 'A1'],
            'Имя приемника': ['Визуал1', 'Визуал2', 'Столбец1'],
            'Тип связи': ['использует', 'использует', 'содержит'],
            'Контекст': ['', '', '']
        }),
        'visuals': pd.DataFrame({
            'ID визуала': ['VIS:1', 'VIS:2'],
            'Тип визуала': ['Карточка', 'График']
        }),
        'excel_path': '/path/to/excel.xlsx',  # строка, не DataFrame
        'columns': None,  # None значение
        'sources': pd.DataFrame(),  # пустой DataFrame
    }
    
    # Устанавливаем current_dfs
    window.current_dfs = test_dfs
    # Сохраняем оригинальные данные (это произойдет при вызове _apply_lineage_focus)
    window._original_dfs = test_dfs.copy()
    
    # Настраиваем comboBox с тестовым значением
    window.lineage_focus_combo.addItem("Мера | M2 | Мера2", userData="M2")
    window.lineage_focus_combo.setCurrentIndex(0)
    
    print("Тестовые данные установлены.")
    print("Ключи в original_dfs:", list(window._original_dfs.keys()))
    for k, v in window._original_dfs.items():
        print(f"  {k}: {type(v).__name__}, значение: {v if not isinstance(v, pd.DataFrame) else 'DataFrame'}")

    # Вызываем метод (должен пройти без AttributeError)
    try:
        window._apply_lineage_focus()
        print("Метод _apply_lineage_focus выполнен без ошибок.")
        print("Результат фильтрации:")
        for key, df in window.current_dfs.items():
            if isinstance(df, pd.DataFrame):
                print(f"  {key}: {len(df)} строк")
            else:
                print(f"  {key}: {type(df).__name__} = {df}")
    except Exception as e:
        print(f"Ошибка: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    # Проверяем, что строковые значения сохранены
    if window.current_dfs.get('excel_path') == '/path/to/excel.xlsx':
        print("Строковое значение 'excel_path' сохранено.")
    else:
        print("Строковое значение потерялось.")
    
    # Проверяем фильтрацию мер
    measures_df = window.current_dfs.get('measures')
    if isinstance(measures_df, pd.DataFrame):
        if len(measures_df) == 1 and measures_df.iloc[0]['№ меры'] == 'M2':
            print("Меры отфильтрованы правильно.")
        else:
            print(f"Меры отфильтрованы неверно: {len(measures_df)} строк")
    else:
        print("Меры не DataFrame.")
    
    # Проверяем фильтрацию lineage
    lineage_df = window.current_dfs.get('lineage')
    if isinstance(lineage_df, pd.DataFrame):
        # Должны остаться строки, где ID источника или приемника = M2
        expected_ids = {'M2'}
        actual_ids = set(lineage_df['ID источника'].tolist() + lineage_df['ID приемника'].tolist())
        if actual_ids.issubset(expected_ids):
            print("Lineage отфильтрован правильно.")
        else:
            print(f"Lineage содержит неожиданные ID: {actual_ids}")
    else:
        print("Lineage не DataFrame.")
    
    print("Тест завершен.")
    sys.exit(0)

if __name__ == '__main__':
    test_filter_with_string_value()
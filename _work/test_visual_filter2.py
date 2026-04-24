#!/usr/bin/env python3
"""Тест фильтрации визуальных элементов."""
import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication

sys.path.insert(0, '.')

from pbi_modules.gui import PBITAnalyzerMainWindow

def test_visual_filter():
    app = QApplication(sys.argv)
    window = PBITAnalyzerMainWindow()
    
    # Создаем тестовые данные с lineage, связывающим M2 и VIS:2
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
            'ID визуала': ['VIS_ID :1', 'VIS_ID :2', 'VIS_ID :3'],
            'Тип визуала': ['Карточка', 'График', 'Таблица'],
            'Имя листа': ['Лист1', 'Лист2', 'Лист3']
        }),
        'excel_path': '/path/to/excel.xlsx',
        'columns': None,
        'sources': pd.DataFrame(),
    }
    
    window.current_dfs = test_dfs
    window._original_dfs = test_dfs.copy()
    
    # Настраиваем comboBox с фокусом на меру M2
    window.lineage_focus_combo.addItem("Мера | M2 | Мера2", userData="M2")
    window.lineage_focus_combo.setCurrentIndex(0)
    
    print("Тестовые данные установлены.")
    print("Визуалы до фильтрации:", len(window._original_dfs['visuals']), "шт.")
    
    # Вызываем фильтрацию
    try:
        window._apply_lineage_focus()
        print("Фильтрация выполнена успешно.")
    except Exception as e:
        print(f"Ошибка при фильтрации: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    # Проверяем результаты
    visuals_df = window.current_dfs.get('visuals')
    if isinstance(visuals_df, pd.DataFrame):
        print(f"Визуалы после фильтрации: {len(visuals_df)} строк")
        if len(visuals_df) == 1:
            visual_id = visuals_df.iloc[0]['ID визуала']
            if visual_id == 'VIS_ID :2':
                print("Визуал VIS_ID :2 правильно отфильтрован.")
            else:
                print(f"Ожидался VIS_ID :2, получен {visual_id}")
        else:
            print("Ожидалась 1 строка, получено", len(visuals_df))
            print(visuals_df)
    else:
        print("Визуалы не DataFrame.")
    
    # Проверяем меры
    measures_df = window.current_dfs.get('measures')
    if isinstance(measures_df, pd.DataFrame) and len(measures_df) == 1 and measures_df.iloc[0]['№ меры'] == 'M2':
        print("Меры отфильтрованы правильно.")
    else:
        print("Меры отфильтрованы неверно.")
    
    # Проверяем lineage
    lineage_df = window.current_dfs.get('lineage')
    if isinstance(lineage_df, pd.DataFrame):
        print(f"Lineage после фильтрации: {len(lineage_df)} строк")
        # Должна быть одна строка с M2 -> VIS:2
        if len(lineage_df) == 1:
            src = lineage_df.iloc[0]['ID источника']
            dst = lineage_df.iloc[0]['ID приемника']
            if src == 'M2' and dst == 'VIS:2':
                print("Lineage отфильтрован правильно.")
            else:
                print(f"Lineage содержит не ту связь: {src} -> {dst}")
        else:
            print("Lineage содержит больше строк, чем ожидалось.")
    
    # Проверяем активность кнопки очистки фильтра (должна быть активна)
    if window.btn_clear_filter.isEnabled():
        print("Кнопка 'Очистить фильтр' активна (ожидаемо).")
    else:
        print("Кнопка 'Очистить фильтр' неактивна (неожиданно).")
    
    print("Тест завершен.")
    sys.exit(0)

if __name__ == '__main__':
    test_visual_filter()
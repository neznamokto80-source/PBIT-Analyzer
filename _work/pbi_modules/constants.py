"""Общие константы приложения: прочерк, пороги, лимиты отображения."""

# Значение «нет данных» во всех таблицах и отчётах.
EMPTY = "—"

# Таблицы lineage больше этого числа строк показываются в облегчённом режиме.
LINEAGE_HEAVY_TABLE_ROWS_THRESHOLD = 1200

# Сколько символов ячейки показывать в GUI (полное значение — в tooltip).
MAX_CELL_CHARS = 200

# PBIX крупнее этого размера читается через on_disk (экономия RAM).
DEFAULT_ON_DISK_THRESHOLD_MB = 500

__all__ = [
    "EMPTY",
    "LINEAGE_HEAVY_TABLE_ROWS_THRESHOLD",
    "MAX_CELL_CHARS",
    "DEFAULT_ON_DISK_THRESHOLD_MB",
]

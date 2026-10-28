"""Чтение JSON из архива Power BI (UTF-16 с BOM и без, fallback на UTF-8)."""
import json


def load_json_utf16(file_path: str) -> dict:
    encodings = ("utf-16-le", "utf-16", "utf-8-sig", "utf-8")
    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as file:
                return json.load(file)
        except (UnicodeError, json.JSONDecodeError):
            # UnicodeError (а не только UnicodeDecodeError): utf-16 без BOM
            # поднимает базовый UnicodeError, и без этого падал fallback на utf-8.
            continue
    raise ValueError(f"Не удалось прочитать JSON-файл с поддерживаемыми кодировками: {file_path}")

__all__ = ["load_json_utf16"]

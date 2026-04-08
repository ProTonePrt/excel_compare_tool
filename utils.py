import re
from typing import Optional

import pandas as pd


# Возможные названия колонки с идентификатором ТС
COLUMN_KEYWORDS = [
    "гос номер",
    "госномер",
    "гос. номер",
    "номер",
    "рег. номер",
    "регистрационный номер",
    "вин",
    "vin",
    "vin номер",
    "vin-код",
    "номер тс",
    "авто",
    "машина",
    "license",
    "plate",
    "registration",
]


def normalize_text(value: str) -> str:
    """Нормализует название колонки для более устойчивого сравнения."""
    text = str(value).strip().lower()
    text = re.sub(r"[\s._\-]+", " ", text)
    return text


def find_plate_column(df: pd.DataFrame) -> Optional[str]:
    """
    Ищет первую подходящую колонку с гос. номером или VIN.
    Возвращает исходное имя колонки или None.
    """
    if df is None:
        return None

    normalized_keywords = [normalize_text(item) for item in COLUMN_KEYWORDS]

    for col in df.columns:
        col_norm = normalize_text(col)
        for keyword in normalized_keywords:
            # Поддерживаем как точное совпадение, так и вхождение.
            if col_norm == keyword or keyword in col_norm:
                return col
    return None


def normalize_plate(value) -> Optional[str]:
    """
    Нормализация значения гос. номера / VIN.
    Возвращает нормализованный текст или None, если значение невалидно.
    """
    if pd.isna(value):
        return None

    text = str(value).strip().upper()
    # Удаляем пробелы, дефисы, точки, слеши и обратные слеши.
    text = re.sub(r"[\s\-./\\]", "", text)

    # Короткие значения считаем мусором.
    if len(text) < 6:
        return None

    return text

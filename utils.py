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

    # 1) Сначала ищем по названию колонки.
    for col in df.columns:
        col_norm = normalize_text(col)
        for keyword in normalized_keywords:
            # Поддерживаем как точное совпадение, так и вхождение.
            if col_norm == keyword or keyword in col_norm:
                return col

    # 2) Если по названию не нашли, пробуем определить по содержимому.
    # Это помогает, когда шапка нестандартная: "ТС", "ID", "Колонка 5" и т.д.
    best_col = None
    best_score = 0

    for col in df.columns:
        series = df[col].dropna()
        if series.empty:
            continue

        # Ограничиваем объём анализа для скорости.
        sample = series.head(80).tolist()
        valid_count = 0

        for value in sample:
            normalized = normalize_plate(value)
            # Валидный идентификатор и в нём есть хотя бы одна цифра.
            if normalized and any(ch.isdigit() for ch in normalized):
                valid_count += 1

        if valid_count > best_score:
            best_score = valid_count
            best_col = col

    # Порог, чтобы не схватить случайный текстовый столбец.
    if best_col is not None and best_score >= 3:
        return best_col

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

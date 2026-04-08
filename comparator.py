from pathlib import Path
from typing import Dict, Optional

import pandas as pd

from utils import find_plate_column, normalize_plate


class ExcelComparator:
    """Класс для загрузки и сравнения Excel-файлов."""

    @staticmethod
    def _to_sorted_strings(values) -> list:
        """
        Безопасно приводит значения к строкам и сортирует.
        Защищает от смешанных типов (str/float и т.д.).
        """
        prepared = []
        for item in values:
            if item is None:
                continue
            prepared.append(str(item))
        return sorted(prepared)

    def load_file(self, path: str) -> pd.DataFrame:
        """Загружает Excel-файл в DataFrame с проверкой типовых ошибок."""
        file_path = Path(path)

        if not file_path.exists():
            raise FileNotFoundError(f"Файл не найден: {path}")

        try:
            df = pd.read_excel(file_path)
        except PermissionError as exc:
            raise PermissionError("Не удалось открыть файл. Закройте файл в Excel и повторите.") from exc
        except ValueError as exc:
            raise ValueError("Файл повреждён или имеет неподдерживаемый формат Excel.") from exc
        except Exception as exc:
            if "format cannot be determined" in str(exc).lower():
                raise ValueError("Не удалось определить формат Excel. Проверьте расширение файла.") from exc
            raise Exception(f"Ошибка чтения файла: {exc}") from exc

        if df.empty:
            raise ValueError("Файл пустой. Добавьте данные и повторите попытку.")

        return df

    def compare(
        self,
        old_df: pd.DataFrame,
        new_df: pd.DataFrame,
        col_old: Optional[str] = None,
        col_new: Optional[str] = None,
    ) -> Dict[str, object]:
        """
        Сравнивает два DataFrame по колонке с идентификатором ТС.
        Возвращает словарь с добавленными, удалёнными и неизменными значениями.
        """
        selected_old_col = col_old or find_plate_column(old_df)
        selected_new_col = col_new or find_plate_column(new_df)

        if not selected_old_col or not selected_new_col:
            raise ValueError(
                "Не удалось найти колонку с гос. номером или VIN.\n"
                "Проверьте названия столбцов, например: 'Гос номер', 'VIN', 'Рег. номер'."
            )

        old_set = set()
        for normalized in old_df[selected_old_col].apply(normalize_plate).tolist():
            if normalized:
                old_set.add(str(normalized))

        new_set = set()
        for normalized in new_df[selected_new_col].apply(normalize_plate).tolist():
            if normalized:
                new_set.add(str(normalized))

        added = self._to_sorted_strings(new_set - old_set)
        removed = self._to_sorted_strings(old_set - new_set)
        unchanged = self._to_sorted_strings(old_set & new_set)

        return {
            "col_old": selected_old_col,
            "col_new": selected_new_col,
            "added": added,
            "removed": removed,
            "unchanged": unchanged,
            "old_total_valid": len(old_set),
            "new_total_valid": len(new_set),
        }

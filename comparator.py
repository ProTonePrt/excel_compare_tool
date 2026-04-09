from pathlib import Path
from typing import Dict, Optional
import json
import time
import uuid
from collections import Counter
import hashlib

import pandas as pd

from utils import find_plate_column, normalize_plate

DEBUG_LOG_PATH = "debug-fff107.log"
DEBUG_SESSION_ID = "fff107"


def _debug_log(hypothesis_id: str, location: str, message: str, data: dict):
    # region agent log
    payload = {
        "sessionId": DEBUG_SESSION_ID,
        "runId": "initial",
        "hypothesisId": hypothesis_id,
        "location": location,
        "message": message,
        "data": data,
        "timestamp": int(time.time() * 1000),
        "id": f"log_{int(time.time() * 1000)}_{uuid.uuid4().hex[:8]}",
    }
    with open(DEBUG_LOG_PATH, "a", encoding="utf-8") as fh:
        fh.write(json.dumps(payload, ensure_ascii=False) + "\n")
    # endregion


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
            xls = pd.ExcelFile(file_path)
            sheet_names = [str(name) for name in xls.sheet_names]
            # region agent log
            _debug_log(
                "H6",
                "comparator.py:load_file",
                "Workbook sheet list",
                {"path": str(file_path), "sheetNames": sheet_names, "sheetCount": int(len(sheet_names))},
            )
            # endregion
            df = pd.read_excel(file_path)
            # region agent log
            _debug_log(
                "H6",
                "comparator.py:load_file",
                "Loaded default sheet",
                {"path": str(file_path), "shape": [int(df.shape[0]), int(df.shape[1])]},
            )
            # endregion
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
        _debug_log(
            "H1",
            "comparator.py:compare",
            "Selected columns",
            {
                "selectedOldCol": str(selected_old_col) if selected_old_col else None,
                "selectedNewCol": str(selected_new_col) if selected_new_col else None,
                "oldRows": int(len(old_df)),
                "newRows": int(len(new_df)),
            },
        )

        if not selected_old_col or not selected_new_col:
            raise ValueError(
                "Не удалось найти колонку с гос. номером или VIN.\n"
                "Проверьте названия столбцов, например: 'Гос номер', 'VIN', 'Рег. номер'."
            )

        old_set = set()
        old_non_empty = 0
        for normalized in old_df[selected_old_col].apply(normalize_plate).tolist():
            if normalized:
                old_non_empty += 1
                old_set.add(str(normalized))

        new_set = set()
        new_non_empty = 0
        for normalized in new_df[selected_new_col].apply(normalize_plate).tolist():
            if normalized:
                new_non_empty += 1
                new_set.add(str(normalized))

        _debug_log(
            "H2",
            "comparator.py:compare",
            "Normalized stats",
            {
                "oldNonEmptyCount": int(old_non_empty),
                "newNonEmptyCount": int(new_non_empty),
                "oldUnique": int(len(old_set)),
                "newUnique": int(len(new_set)),
                "oldSample": list(sorted(old_set))[:5],
                "newSample": list(sorted(new_set))[:5],
            },
        )
        old_counter = Counter(old_df[selected_old_col].apply(normalize_plate).dropna().astype(str).tolist())
        new_counter = Counter(new_df[selected_new_col].apply(normalize_plate).dropna().astype(str).tolist())
        old_fingerprint = hashlib.sha256("|".join(sorted(old_counter.elements())).encode("utf-8")).hexdigest()
        new_fingerprint = hashlib.sha256("|".join(sorted(new_counter.elements())).encode("utf-8")).hexdigest()
        changed_counts = []
        all_keys = set(old_counter.keys()) | set(new_counter.keys())
        for key in sorted(all_keys):
            if old_counter.get(key, 0) != new_counter.get(key, 0):
                changed_counts.append(
                    {"id": key, "oldCount": int(old_counter.get(key, 0)), "newCount": int(new_counter.get(key, 0))}
                )
            if len(changed_counts) >= 5:
                break
        _debug_log(
            "H5",
            "comparator.py:compare",
            "Multiset diagnostics",
            {
                "oldFingerprint": old_fingerprint,
                "newFingerprint": new_fingerprint,
                "sameFingerprint": old_fingerprint == new_fingerprint,
                "topCountChanges": changed_counts,
            },
        )

        added = self._to_sorted_strings(new_set - old_set)
        removed = self._to_sorted_strings(old_set - new_set)
        unchanged = self._to_sorted_strings(old_set & new_set)
        _debug_log(
            "H3",
            "comparator.py:compare",
            "Diff result",
            {
                "addedCount": int(len(added)),
                "removedCount": int(len(removed)),
                "unchangedCount": int(len(unchanged)),
                "addedSample": added[:5],
                "removedSample": removed[:5],
            },
        )

        return {
            "col_old": selected_old_col,
            "col_new": selected_new_col,
            "added": added,
            "removed": removed,
            "unchanged": unchanged,
            "old_total_valid": len(old_set),
            "new_total_valid": len(new_set),
        }

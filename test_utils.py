import unittest

import pandas as pd

from utils import find_plate_column, normalize_plate


class TestUtils(unittest.TestCase):
    def test_normalize_plate_basic(self):
        self.assertEqual(normalize_plate("с 136 ам 198"), "С136АМ198")
        self.assertEqual(normalize_plate("С-136-АМ-198"), "С136АМ198")
        self.assertEqual(normalize_plate("XTA219000J0123456"), "XTA219000J0123456")

    def test_normalize_plate_invalid(self):
        self.assertIsNone(normalize_plate(None))
        self.assertIsNone(normalize_plate("123"))
        self.assertIsNone(normalize_plate("  -  "))

    def test_find_plate_column_various_names(self):
        df_old = pd.DataFrame({"Гос номер": ["А123АА77"]})
        df_new = pd.DataFrame({"Госномер": ["А123АА77"]})
        df_vin = pd.DataFrame({"VIN номер": ["XTA219000J0123456"]})

        self.assertEqual(find_plate_column(df_old), "Гос номер")
        self.assertEqual(find_plate_column(df_new), "Госномер")
        self.assertEqual(find_plate_column(df_vin), "VIN номер")

    def test_find_plate_column_not_found(self):
        df = pd.DataFrame({"Дата": ["2026-01-01"], "Комментарий": ["тест"]})
        self.assertIsNone(find_plate_column(df))

    def test_find_plate_column_by_content_fallback(self):
        df = pd.DataFrame(
            {
                "Колонка A": ["текст", "комментарий", "данные"],
                "ID ТС": ["С 136 АМ 198", "К777ТТ197", "XTA219000J0123456"],
            }
        )
        self.assertEqual(find_plate_column(df), "ID ТС")

    def test_find_plate_column_does_not_pick_numeric_column(self):
        df = pd.DataFrame(
            {
                "Дата": ["2026-04-01", "2026-04-02", "2026-04-03"],
                "Сумма": [1000, 2000, 3000],
                "ТС": ["А123ВС77", "К777ТТ197", "XTA219000J0123456"],
            }
        )
        self.assertEqual(find_plate_column(df), "ТС")


if __name__ == "__main__":
    unittest.main()

import unittest

import pandas as pd

from comparator import ExcelComparator


class TestExcelComparator(unittest.TestCase):
    def setUp(self):
        self.comparator = ExcelComparator()

    def test_compare_mixed_formats(self):
        old_df = pd.DataFrame(
            {
                "Гос номер": [
                    "С 136 АМ 198",
                    "К777ТТ197",
                    "XTA219000J0123456",
                    "Н123НО47",  # удалится
                ]
            }
        )
        new_df = pd.DataFrame(
            {
                "Госномер": [
                    "С-136-АМ-198",
                    "К 777 ТТ 197",
                    "XTA219000J0123456",
                    "Р999РР77",  # добавится
                ]
            }
        )

        result = self.comparator.compare(old_df, new_df)

        self.assertEqual(result["col_old"], "Гос номер")
        self.assertEqual(result["col_new"], "Госномер")
        self.assertEqual(result["added"], ["Р999РР77"])
        self.assertEqual(result["removed"], ["Н123НО47"])
        self.assertEqual(sorted(result["unchanged"]), ["XTA219000J0123456", "К777ТТ197", "С136АМ198"])

    def test_compare_raises_if_column_not_found(self):
        old_df = pd.DataFrame({"Дата": ["2026-01-01"]})
        new_df = pd.DataFrame({"Комментарий": ["тест"]})

        with self.assertRaises(ValueError):
            self.comparator.compare(old_df, new_df)

    def test_compare_with_mixed_raw_types_does_not_crash(self):
        old_df = pd.DataFrame({"ID": ["A123BC77", 123456.0, "XTA219000J0123456"]})
        new_df = pd.DataFrame({"ID": ["A123BC77", "123456.0", "XTA219000J0123456"]})

        result = self.comparator.compare(old_df, new_df, col_old="ID", col_new="ID")

        self.assertIsInstance(result["added"], list)
        self.assertIsInstance(result["removed"], list)
        self.assertIsInstance(result["unchanged"], list)


if __name__ == "__main__":
    unittest.main()

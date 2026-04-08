from pathlib import Path

import pandas as pd


def main():
    base_dir = Path(__file__).resolve().parent
    old_path = base_dir / "old_vehicles.xlsx"
    new_path = base_dir / "new_vehicles.xlsx"

    # 10 одинаковых + 2 удалённых (есть только в старом)
    old_ids = [
        "С 136 АМ 198",
        "К777ТТ197",
        "А065ПР178",
        "М 001 ММ 77",
        "Е500КХ750",
        "Х111ХХ99",
        "Р444ОР116",
        "У222УУ16",
        "В333ВВ190",
        "Т888ТТ178",
        "Н123НО47",      # будет удалён
        "О456ОО47",      # будет удалён
        "XTA219000J0123456",  # VIN (общий)
    ]

    # 10 одинаковых + 3 новых (есть только в новом), форматы разные
    new_ids = [
        "С-136-АМ-198",
        "К 777 ТТ 197",
        "А065ПР178",
        "м001мм77",
        "Е 500 КХ 750",
        "Х111ХХ99",
        "Р-444-ОР-116",
        "У222УУ16",
        "В333ВВ190",
        "Т 888 ТТ 178",
        "Р999РР77",      # новый
        "L 234 LL 77",   # новый
        "ZFA22300005555555",  # новый VIN
        "XTA219000J0123456",  # VIN (общий)
    ]

    old_df = pd.DataFrame({"Гос номер": old_ids})
    new_df = pd.DataFrame({"Госномер": new_ids})

    old_df.to_excel(old_path, index=False)
    new_df.to_excel(new_path, index=False)

    print(f"Создано: {old_path}")
    print(f"Создано: {new_path}")


if __name__ == "__main__":
    main()

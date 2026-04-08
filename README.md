# Excel Compare Tool

Простое десктоп-приложение на Python (`tkinter`) для сравнения двух Excel-файлов со списками транспортных средств по гос. номерам или VIN.

## Возможности

- Выбор старого и нового Excel-файла (`.xlsx`, `.xls`)
- Автоматический поиск колонки с идентификатором ТС (гос. номер / VIN)
- Нормализация значений (удаление пробелов, дефисов, точек, слешей, верхний регистр)
- Определение:
  - какие ТС добавились
  - какие ТС выбыли
  - какие остались без изменений
- Вывод отчёта в приложении и сохранение в `.txt`

## Структура проекта

```text
excel_compare_tool/
├── main.py
├── comparator.py
├── utils.py
├── requirements.txt
├── README.md
└── test_data/
    ├── old_vehicles.xlsx
    └── new_vehicles.xlsx
```

## Установка Python

Скачайте и установите Python 3.9+ с официального сайта: [python.org](https://www.python.org/downloads/)

## Установка и запуск

1. Перейдите в папку проекта:

```bash
cd excel_compare_tool
```

2. Создайте виртуальное окружение:

```bash
python -m venv venv
```

3. Активируйте окружение (Windows):

```bash
venv\Scripts\activate
```

4. Установите зависимости:

```bash
pip install -r requirements.txt
```

5. Запустите приложение:

```bash
python main.py
```

## Сборка в .exe (PyInstaller)

Установите PyInstaller:

```bash
pip install pyinstaller
```

Соберите приложение:

```bash
pyinstaller --onefile --windowed --name "ExcelCompare" main.py
```

По вашему запросу также подходит команда:

```bash
pyinstaller --onefile --windowed main.py
```

### Быстрая сборка через bat

Можно запустить готовый скрипт:

```bat
build_exe.bat
```

После успешной сборки получите файл:

```text
dist\ExcelCompare.exe
```

## Создание установщика (installer .exe)

Если нужен установщик с мастером установки (а не только portable `.exe`), используйте Inno Setup:

1. Установите Inno Setup: [jrsoftware.org/isinfo.php](https://jrsoftware.org/isinfo.php)
2. Убедитесь, что уже собран `dist\ExcelCompare.exe`
3. Откройте файл `installer.iss` в Inno Setup
4. Нажмите **Compile**

Результат:

```text
dist\ExcelCompareInstaller.exe
```

## Portable-версия (без установки Python и программы)

Для portable-варианта (запуск сразу с флешки/папки) используйте:

```bat
build_portable.bat
```

Скрипт создаёт:

```text
dist\ExcelComparePortable.exe
dist\ExcelComparePortable.zip
portable\
```

Как использовать:

1. Передайте пользователю `ExcelComparePortable.zip` или `ExcelComparePortable.exe`
2. На целевом ПК Python устанавливать не нужно
3. Для запуска достаточно открыть `ExcelComparePortable.exe`

## Сборка в облаке (GitHub Actions)

Если на рабочем ПК запрещено устанавливать Python, используйте облачную сборку:

1. Загрузите проект в GitHub-репозиторий
2. Откройте вкладку **Actions**
3. Выберите workflow **Build Portable EXE**
4. Нажмите **Run workflow**
5. После завершения откройте run и скачайте артефакты:
   - `ExcelComparePortable-exe`
   - `ExcelComparePortable-zip`

Workflow уже добавлен в проект:

```text
.github/workflows/build-portable.yml
```

## Тестовые данные

В папке `test_data` находятся примерные файлы:

- `old_vehicles.xlsx` — старый список
- `new_vehicles.xlsx` — новый список
- `generate_test_data.py` — генератор тестовых файлов

Сценарий включает:

- 10 одинаковых идентификаторов
- 3 новых (поступили)
- 2 удалённых (выбыли)
- разные форматы записи (пробелы, дефисы, разный регистр)
- присутствует VIN

Если файлов `old_vehicles.xlsx` и `new_vehicles.xlsx` ещё нет, создайте их командой:

```bash
python test_data/generate_test_data.py
```

## Запуск тестов

Тесты написаны на стандартном `unittest` (без дополнительных зависимостей).

Запуск всех тестов:

```bash
python -m unittest discover -s tests -p "test_*.py"
```

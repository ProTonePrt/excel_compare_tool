@echo off
setlocal

set "PYTHON_CMD="

python -c "import sys; print(sys.version)" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=python"

if not defined PYTHON_CMD (
    if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python312\python.exe"
)
if not defined PYTHON_CMD (
    if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python311\python.exe"
)
if not defined PYTHON_CMD (
    if exist "%LocalAppData%\Programs\Python\Python310\python.exe" set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python310\python.exe"
)
if not defined PYTHON_CMD (
    if exist "%LocalAppData%\Programs\Python\Python39\python.exe" set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python39\python.exe"
)

if not defined PYTHON_CMD (
    echo Ошибка: Python не найден.
    echo Установите Python 3.9+ с https://www.python.org/downloads/
    echo и включите опцию "Add Python to PATH".
    pause
    exit /b 1
)

echo [1/4] Создание виртуального окружения...
"%PYTHON_CMD%" -m venv venv
if errorlevel 1 (
    echo Ошибка: не удалось создать venv. Убедитесь, что Python установлен и добавлен в PATH.
    pause
    exit /b 1
)

echo [2/4] Активация окружения...
set "VENV_PYTHON=venv\Scripts\python.exe"
if not exist "%VENV_PYTHON%" (
    echo Ошибка: не найден %VENV_PYTHON%
    pause
    exit /b 1
)

echo [3/4] Установка зависимостей...
"%VENV_PYTHON%" -m pip install --upgrade pip
"%VENV_PYTHON%" -m pip install -r requirements.txt
"%VENV_PYTHON%" -m pip install pyinstaller
if errorlevel 1 (
    echo Ошибка: не удалось установить зависимости.
    pause
    exit /b 1
)

echo [4/4] Сборка EXE...
"%VENV_PYTHON%" -m PyInstaller --onefile --windowed --name "ExcelCompare" main.py
if errorlevel 1 (
    echo Ошибка: сборка завершилась с ошибкой.
    pause
    exit /b 1
)

echo.
echo Готово! EXE-файл: dist\ExcelCompare.exe
pause
endlocal

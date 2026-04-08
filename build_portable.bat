@echo off
cd /d "%~dp0"
call :run > build_portable.log 2>&1
start notepad build_portable.log
pause
exit /b

:run
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

echo [1/6] Создание виртуального окружения...
"%PYTHON_CMD%" -m venv venv
if errorlevel 1 (
    echo Ошибка: не удалось создать venv. Убедитесь, что Python установлен и добавлен в PATH.
    pause
    exit /b 1
)

echo [2/6] Активация окружения...
set "VENV_PYTHON=venv\Scripts\python.exe"
if not exist "%VENV_PYTHON%" (
    echo Ошибка: не найден %VENV_PYTHON%
    pause
    exit /b 1
)

echo [3/6] Установка зависимостей...
"%VENV_PYTHON%" -m pip install --upgrade pip
"%VENV_PYTHON%" -m pip install -r requirements.txt
"%VENV_PYTHON%" -m pip install pyinstaller
if errorlevel 1 (
    echo Ошибка: не удалось установить зависимости.
    pause
    exit /b 1
)

echo [4/6] Сборка portable EXE...
"%VENV_PYTHON%" -m PyInstaller --onefile --windowed --name "ExcelComparePortable" main.py
if errorlevel 1 (
    echo Ошибка: сборка завершилась с ошибкой.
    pause
    exit /b 1
)

echo [5/6] Подготовка portable-папки...
if exist portable rmdir /s /q portable
mkdir portable
copy dist\ExcelComparePortable.exe portable\ExcelComparePortable.exe >nul
if exist test_data xcopy test_data portable\test_data\ /e /i /y >nul

echo [6/6] Упаковка в ZIP...
powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Test-Path 'dist\\ExcelComparePortable.zip') { Remove-Item 'dist\\ExcelComparePortable.zip' -Force }; Compress-Archive -Path 'portable\\*' -DestinationPath 'dist\\ExcelComparePortable.zip' -Force"
if errorlevel 1 (
    echo Предупреждение: EXE создан, но ZIP не удалось собрать.
    echo Проверьте папку portable\
    pause
    exit /b 0
)

echo.
echo Готово!
echo Portable EXE: dist\ExcelComparePortable.exe
echo Portable ZIP: dist\ExcelComparePortable.zip
echo Папка для ручного копирования: portable\
pause
endlocal

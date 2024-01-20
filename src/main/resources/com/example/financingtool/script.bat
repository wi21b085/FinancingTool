@echo off
set arg1=%1
set arg2=%2

IF NOT EXIST "src\\main\\resources\\com\\example\\financingtool\\pre.txt" (
    echo Prerequisites not found
    echo ...
    echo Installing Python...

    where python > nul 2>&1

    if %errorlevel% equ 0 (
        python --version
        echo "Python is installed"
    ) else (
        echo "Python is not installed"
        winget install -e --id Python.Python.3.11 --no-upgrade --accept-package-agreements --accept-source-agreements
        REM winget install python --no-upgrade --accept-package-agreements --accept-source-agreements
    )
    echo ...
    echo Upgrading pip...
    python -m pip install --upgrade pip
    echo ...
    echo Installing Selenium...
    python -c "import selenium" > nul 2>&1

    if %errorlevel% equ 0 (
        echo "Selenium is installed"
    ) else (
        echo "Selenium is not installed"
        pip install --upgrade selenium
    )
    echo ...
    echo Installing PIL...
    python -c "import PIL" > nul 2>&1

    if %errorlevel% equ 0 (
        echo "Pillow (or PIL) is installed"
    ) else (
        echo "Pillow (or PIL) is not installed"
        pip install --upgrade pillow
    )
    echo ...
    echo Installing colorama...
    python -c "import colorama" > nul 2>&1

    if %errorlevel% equ 0 (
        echo "Colorama is installed"
    ) else (
        echo "Colorama is not installed"
        pip install --upgrade colorama
    )
    echo ...
    echo Prerequisites installed> src\\main\\resources\\com\\example\\financingtool\\pre.txt
)
echo Taking screenshot...
python src\\main\\resources\\com\\example\\financingtool\\%arg1% %arg2%

exit
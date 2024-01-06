@echo off
set arg1=%1

IF NOT EXIST "src\\main\\resources\\com\\example\\financingtool\\pre.txt" (
    echo Prerequisites not found
    echo ...
    echo Installing Python...
    winget install -e --id Python.Python.3.12 --no-upgrade --accept-package-agreements --accept-source-agreements
    ::winget install python --no-upgrade --accept-package-agreements --accept-source-agreements
    echo ...
    echo Upgrading pip...
    python -m pip install --upgrade pip
    echo ...
    echo Installing Selenium...
    pip install --upgrade selenium
    echo ...
    echo Installing PIL...
    pip install --upgrade pillow
    echo ...
    echo Prerequisites installed> src\\main\\resources\\com\\example\\financingtool\\pre.txt
)
echo Taking screenshot...
python src\\main\\resources\\com\\example\\financingtool\\widmung.py %arg1%

exit
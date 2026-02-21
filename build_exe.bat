@echo off
echo Сборка exe файла...
pyinstaller --onefile --windowed --name "SearchSet Builder" search_set.py
echo Готово! exe файл находится в папке dist
pause

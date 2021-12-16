
md build

copy ExcelGradingRubricCreator.py build/
cd build
pyinstaller --onefile ExcelGradingRubricCreator.py

md build

copy ExcelGradingRubricCreator.py build/
cd build
pyinstaller --onefile --noconsole ExcelGradingRubricCreator.py
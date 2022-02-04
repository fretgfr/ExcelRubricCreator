
md build

copy ExcelGradingRubricCreator.py build/
cd build
pyinstaller --onefile -upx-dir=C:\upx ExcelGradingRubricCreator.py
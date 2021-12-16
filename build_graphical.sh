
mkdir build 2>/dev/null

cp ExcelGradingRubricCreator.py build/
cd build
pyinstaller --windowed --onefile ExcelGradingRubricCreator.py

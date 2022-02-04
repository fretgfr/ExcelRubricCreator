
md build

copy GraphicalRubricCreator.py build
cd build
pyinstaller -F --upx-dir=C:\upx GraphicalRubricCreator.py
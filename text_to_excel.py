"""
    @File: TextToExcel.py
    @Author: Younghoon Jeong (yhjeong0422@gmail.com)

    This program moves data in text files to an Excel file, grouping
    data files with serial number in the name of data files.

    Required libraries:
        xlwings (pip)
        PyQt5 (pip3)
        pywin32 (Internet download required)

    This file functions as a main function.
"""

__version__ = '1.0.1'

from PyQt5.QtWidgets import *
from gui.messages import ErrorMessage
from gui.mainwindow import MainWindow
from xlwings import *
import sys

app = QApplication(sys.argv)
if __name__ == "__main__":
    try:
        if int(sys.version[0]) != 3 and float(sys.version[:3]) < 3.6:
            raise Exception
    except:
        errorMessage = ErrorMessage()
        errorMessage.setText("<nobr>The version of Python must be 3.6 or higher.</nobr>")
        sys.exit(app.exec_())

excelApp = App()
excelApp.visible = False
excelApp.display_alerts = False
for book in excelApp.books:
    book.close()
win = MainWindow()
win.resize(3000, 1000)
win.show()
try:
    app.exec_()
except Exception as e:
    print(e)
finally:
    for book in excelApp.books:
        book.close()
    excelApp.quit()


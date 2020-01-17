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
import sys
from basic.file.files import ExcelFile
from win32api import GetSystemMetrics

app = QApplication(sys.argv)
if __name__ == "__main__":
  try:
    if int(sys.version[0]) != 3 and float(sys.version[:3]) < 3.6:
      raise Exception
  except:
    errorMessage = ErrorMessage()
    errorMessage.setText("<nobr>The version of Python must be 3.6 or higher.</nobr>")
    sys.exit(app.exec_())

ExcelFile.open_excel_app()
win = MainWindow()
win.resize(GetSystemMetrics(0) * 3 / 4.0, GetSystemMetrics(1) * 3 / 5.0)
win.show()
try:
  app.exec_()
except Exception as e:
  print(e)
finally:
  ExcelFile.close_excel_app()

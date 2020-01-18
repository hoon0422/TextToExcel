"""
    This module has a class for GUI to show a widget to decide a template to add data files.
"""

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from basic.file.files import ExcelFile
from gui.messages import WaitingMessage, ErrorMessage
import xlwings as xw
import sys, gui


class WgtTemplate(QFrame):
  """ This class represents a widget to decide a template to add data files.

      Attributes:
          _bt_select: a button to select a template.
          _lbl_template: a label that shows a template name.
          _template: an Excel file for a template.
  """
  template_changed = pyqtSignal(ExcelFile, name="template_changed")

  def __init__(self):
    super(WgtTemplate, self).__init__()
    self._bt_select = QPushButton("Select Template")
    self._lbl_template = QLabel("...")
    self._template: ExcelFile = None

    h_layout = QHBoxLayout()
    h_layout.addWidget(self._bt_select)
    h_layout.addSpacerItem(QSpacerItem(70, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
    v_layout = QVBoxLayout()
    v_layout.addLayout(h_layout)
    v_layout.addWidget(self._lbl_template)

    self.setLayout(v_layout)
    self.setFrameShape(QFrame.StyledPanel)

    self._bt_select.clicked.connect(self.__bt_select_clicked)
    self.template_changed.connect(self.__printing_changed)

    self.setAcceptDrops(True)

  # Getters
  @property
  def template(self):
    return self._template

  def __set_template(self, excel_file_name: str):
    loading_mb = WaitingMessage("Loading Excel file...")
    loading_mb.show()
    try:
      self._template = ExcelFile(excel_file_name)
      for app in xw.apps:
        for book in app.books:
          if book.fullname.upper() == self._template.full_name.upper():
            raise OSError("Another program use the file. Please close the program.")
      self.template_changed.emit(self._template)
      loading_mb.accept()
      self._lbl_template.setText(self._template.name + self._template.file_format)
    except Exception as e:
      error_mb = ErrorMessage("<nobr>Error: " + str(e) + "</nobr>")
      error_mb.show()
      error_mb.exec_()
      loading_mb.reject()

  def __bt_select_clicked(self):
    fname = QFileDialog.getOpenFileName(self, 'Open file', directory=gui.current_directory,
                                        filter="Excel 2007 (*.xlsx)")

    if fname[0] != "":
      gui.current_directory = fname[0][:fname[0].rfind('\\')]
      self.__set_template(fname[0])

  @pyqtSlot(ExcelFile)
  def __printing_changed(self, excel: ExcelFile):
    print("Current template: " + str(excel.name + excel.file_format))

  # Drag - Drop functions to enable to drag and drop an Excel file from OS Explorer.
  def dragEnterEvent(self, event):
    if event.mimeData().hasUrls():
      event.accept()
    else:
      event.ignore()

  def dragMoveEvent(self, event):
    if event.mimeData().hasUrls():
      event.accept()
    else:
      event.ignore()

  def dropEvent(self, event):
    if event.mimeData().hasUrls():
      event.accept()
      file_names = []
      for url in event.mimeData().urls():
        file_names.append(url.toLocalFile())
      if len(file_names) != 1:
        error_mb = ErrorMessage("<nobr>Only one Excel file can be a template</nobr>")
        error_mb.show()
        error_mb.exec_()
        return

      dot_index = file_names[0].rfind('.')
      if dot_index == -1 or file_names[0][dot_index:] != ExcelFile.FORMAT:
        error_mb = ErrorMessage("<nobr>Only Excel file (.xlsx extension) can be template.</nobr>")
        error_mb.show()
        error_mb.exec_()
        return

      self.__set_template(file_names[0])

    else:
      event.ignore()


if __name__ == "__main__":
  app = QApplication(sys.argv)
  window = WgtTemplate()
  window.show()
  app.exec_()

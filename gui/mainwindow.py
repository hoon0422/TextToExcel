"""
    This module has a GUI class for the main window.
"""

from gui.datatable import *
from gui.sheetinfo import *
from gui.template import *
from gui.datalist import *
from gui.dialogs import *
from gui.sheetkeyword import *
from gui.copiedrange import *
from basic.file import check_valid_range
from gui.messages import ErrorMessage


class MainWindow(QMainWindow):
  """
      Main window class

      Attributes:
          _template: WgtTemplate object
          _data_list: WgtDataList object
          _sheet_info: WgtSheetInfo object
          _data_table: WgtDataTable object
          _keyword: WgtKeyword object.
  """

  def __init__(self):
    super(MainWindow, self).__init__()
    self._template = WgtTemplate()
    self._data_list = WgtDataList()
    self._sheet_info = WgtSheetInfo()
    self._data_table = WgtDataTable()
    self._keyword = WgtKeyword()
    self._range = WgtCopiedRange()

    self._data_table.data_table.set_keyword(self._keyword.keyword())

    # layouts
    left_layout = QVBoxLayout()
    left_layout.addWidget(self._template)
    left_layout.addWidget(self._keyword)
    left_layout.addWidget(self._range)
    left_layout.addWidget(self._data_list, stretch=6)
    left_widget = QFrame()
    left_widget.setLayout(left_layout)
    left_widget.setFrameShape(QFrame.StyledPanel)

    splitter = QSplitter(Qt.Horizontal)
    splitter.addWidget(left_widget)
    splitter.addWidget(self._sheet_info)
    splitter.addWidget(self._data_table)
    splitter.setChildrenCollapsible(False)
    splitter.setStretchFactor(0, 1)
    splitter.setStretchFactor(1, 4)
    splitter.setStretchFactor(2, 4)
    splitter.setLineWidth(2)
    splitter.setHandleWidth(2)

    self.setCentralWidget(splitter)

    # menu
    menubar = self.menuBar()
    mn_file = menubar.addMenu('File')

    about_action = mn_file.addAction('About')
    about_action.triggered.connect(self.about_action_triggered)

    # signals - slots
    self._template.template_changed.connect(self._sheet_info.info_table.set_excel_file)
    self._template.template_changed.connect(self._data_table.data_table.set_sheet_list)
    self._data_list.data_list.list_changed.connect(self._data_table.data_table.set_serial_groups)
    self._data_table.bt_set_save_file_names.clicked.connect(self.__bt_set_save_file_names)
    self._keyword.keyword_changed.connect(self._data_table.data_table.set_keyword)

    self.setWindowTitle("Text to Excel")

  def about_action_triggered(self):
    dlg = DlgAboutProgram()
    dlg.show()
    dlg.exec_()

  def __bt_set_save_file_names(self):
    if check_valid_range(self._range.excel_range()):
      DlgSaveFileName(self._template.template,
                      self._data_table.data_table.get_table(),
                      self._sheet_info.info_table.get_sheet_data(),
                      self._range.excel_range()).show()
    else:
      error_mb = ErrorMessage("Not valid range. Please check range again")
      error_mb.show()
      error_mb.exec_()
      return


if __name__ == "__main__":
  app = QApplication(sys.argv)
  window = MainWindow()
  window.resize(3000, 1000)
  window.show()
  app.exec_()

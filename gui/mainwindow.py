"""
    This module has a GUI class for the main window.
"""

from gui.datatable import *
from gui.sheetinfo import *
from gui.template import *
from gui.datalist import *
from gui.dialogs import *
from gui.sheetkeyword import *


class MainWindow(QMainWindow):
  """
      Main window class

      Attributes:
          _template: WgtTemplate object
          _data_liat: WgtDataList object
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

    self._data_table.data_table.set_keyword(self._keyword.keyword())

    # layouts
    leftLayout = QVBoxLayout()
    leftLayout.addWidget(self._template)
    leftLayout.addWidget(self._keyword)
    leftLayout.addWidget(self._data_list, stretch=6)
    leftWidget = QFrame()
    leftWidget.setLayout(leftLayout)
    leftWidget.setFrameShape(QFrame.StyledPanel)

    splitter = QSplitter(Qt.Horizontal)
    splitter.addWidget(leftWidget)
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
    self._data_table.bt_set_save_file_names.clicked.connect(
      lambda: DlgSaveFileName(self._template.template,
                              self._data_table.data_table.get_table(),
                              self._sheet_info.info_table.get_sheet_data()).show())
    self._keyword.keyword_changed.connect(self._data_table.data_table.set_keyword)

    self.setWindowTitle("Text to Excel")

  def about_action_triggered(self):
    dlg = DlgAboutProgram()
    dlg.show()
    dlg.exec_()


if __name__ == "__main__":
  app = QApplication(sys.argv)
  window = MainWindow()
  window.resize(3000, 1000)
  window.show()
  app.exec_()

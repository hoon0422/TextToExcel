"""
    This module has classes for GUI showing a table to decide how to manipulate worksheet.
"""

from basic.sheetdata.sheetdata import *
from basic.file.files import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from itertools import product
import sys


class SearchBar(QWidget):
  """ This class represents a search bar to search worksheets.

      Attributes:
          _edt_search: QLineEdit object for search key.
  """

  def __init__(self):
    super(SearchBar, self).__init__()
    self._edt_search = QLineEdit()

    layout = QFormLayout()
    layout.addRow(QLabel("Search"), self._edt_search)

    self.setLayout(layout)

  @property
  def edt_search(self):
    return self._edt_search

  def keyword(self) -> str:
    return self._edt_search.text()

  @pyqtSlot(ExcelFile)
  def clear_bar(self):
    self._edt_search.setText("")


class SheetInfoTable(QTableWidget):
  """ This class represents a table to decide how to manipulate worksheets.

      Attributes:
          _all_values: a list of widgets to control all the values.
  """

  @classmethod
  def __match_info_with_type(self, info: SheetInfo):
    """
    Returns a proper widget matching with data type.
    :param info: SheetInfo object.
    :return: a proper widget matching with data type.
    :raise TypeError: if there is no proper widget.
    """
    if info.info_type is str:
      return QLineEdit()
    elif info.info_type is bool:
      return QCheckBox()
    else:
      raise TypeError("There is no matched GUI type for " + str(info.info_type))

  def __init__(self):
    super(SheetInfoTable, self).__init__()
    self._all_values = []

    self.setRowCount(len(sheet_infos))
    for i, info in enumerate(sheet_infos):
      self.setVerticalHeaderItem(i, QTableWidgetItem(info.info_name))
    self.__init_all()

    self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
    self.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

  # Getters
  @property
  def info_table(self):
    return self._info_table

  @pyqtSlot(ExcelFile)
  def set_excel_file(self, excel_file: ExcelFile):
    """
    Update the table corresponding to an Excel file.
    :param excel_file: an Excel file.
    """
    self.clearContents()

    # all
    self.__init_all()

    # horizontal header
    with excel_file as excel:
      sheet_names: List[str] = []
      for i, sheet in enumerate(excel.sheets):
        sheet_names.append(sheet.name)

    # cells
    self.setColumnCount(len(sheet_names) + 1)
    self.setHorizontalHeaderLabels(['All'] + sheet_names)
    for r, c in product(range(len(sheet_infos)), range(len(sheet_names))):
      self.setCellWidget(r, c + 1, self.__make_connected_cell(r, c))

  def get_sheet_data(self) -> List[SheetData]:
    """
    Returns data of how to manipulate a worksheet. in the table.
    :return: a list of "SheetData"s
    :raise TypeError: if there is no data type matching with a widget in a cell.
    """
    result: List[SheetData] = []

    for c in range(self.columnCount() - 1):
      s = SheetData(self.horizontalHeaderItem(c + 1).text())
      for r, v in enumerate(sheet_infos):
        wgt = self.cellWidget(r, c + 1)
        if type(wgt) is QLineEdit:
          s[v] = wgt.text()
        elif type(wgt) is QCheckBox:
          s[v] = wgt.isChecked()
        else:
          raise TypeError("There is no matched data type for " + str(type(wgt)))
      result.append(s)

    return result

  def search_sheet(self, keyword: str):
    """
    Searches sheets and show the searched only.
    :param keyword: search key
    """
    for i in range(1, len(self.horizontalHeader())):
      if self.horizontalHeaderItem(i).text().count(keyword) == 0:
        self.setColumnHidden(i, True)
      else:
        self.setColumnHidden(i, False)

  def __init_all(self):
    """ Initializes _all_values and update the first column of the table. """
    self._all_values = []

    self.setColumnCount(1)
    self.setHorizontalHeaderItem(0, QTableWidgetItem("All"))
    for i, info in enumerate(sheet_infos):
      self._all_values.append(SheetInfoTable.__match_info_with_type(info))
      self.setCellWidget(i, 0, self._all_values[i])

  def __make_connected_cell(self, info_index: int, sheet_index: int) -> QWidget:
    """
    Returns a connected widget connected with a widget in _all_values.
    :param info_index: row index of a cell
    :param sheet_index: column index of a cell
    :return: a connected widget.
    """
    wgt = type(self._all_values[info_index])()

    if sheet_infos[info_index].info_type is str:
      self._all_values[info_index].textChanged.connect(
        lambda: wgt.setText(self._all_values[info_index].text())
        if not self.isColumnHidden(sheet_index + 1) else None)
      return wgt
    if sheet_infos[info_index].info_type is bool:
      self._all_values[info_index].stateChanged.connect(
        lambda: wgt.setChecked(self._all_values[info_index].isChecked())
        if not self.isColumnHidden(sheet_index + 1) else None)
      return wgt
    raise TypeError("There is no matched GUI type for " + str(sheet_infos[info_index].info_type))


class WgtSheetInfo(QFrame):
  """ This class represents a widget to decide how to manipulate worksheets.

      Attributes:
          _sheet_search: SearchBar object
          _info_table: SheetInfoTable object.
  """

  def __init__(self):
    super(WgtSheetInfo, self).__init__()
    self._sheet_search = SearchBar()
    self._info_table = SheetInfoTable()

    layout = QVBoxLayout()
    layout.addWidget(self._sheet_search)
    layout.addWidget(self._info_table, 1)

    # for test
    '''test_button = QPushButton("test")
    test_button.clicked.connect(lambda: print(self._info_table.get_sheet_data()))
    layout.addWidget(test_button)'''

    self.setLayout(layout)
    self.setFrameShape(QFrame.StyledPanel)

    self._sheet_search.edt_search.textChanged. \
      connect(lambda: self._info_table.search_sheet(self._sheet_search.keyword()))

  # Getters
  @property
  def sheet_search(self):
    return self._sheet_search

  @property
  def info_table(self):
    return self._info_table


if __name__ == "__main__":
  app = QApplication(sys.argv)
  window = WgtSheetInfo()
  window.info_table.set_excel_file(ExcelFile('C:/Users/yhjeo/Others/test.xlsx'))
  window.show()
  app.exec_()

"""
    This module has classes for GUI showing data table to match data files with serials and worksheets.
"""

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from basic.list2d import Table
from basic.file.files import *
from basic.file import group_data_files
from typing import List
from itertools import product

DataTable = Table[TextFile, str, str]


class DataComboBox(QComboBox):
  """ This class represents a combo box that has a list of data files.

      Attributes:
          _data_list: a list of data files
  """

  def __init__(self):
    super(DataComboBox, self).__init__()
    self._data_list: List[TextFile] = []

  # Getters
  @property
  def data_list(self):
    return self._data_list

  def add(self, data: TextFile):
    """ Adds a data file to the list.
    :param data: a data file.
    """
    if self._data_list.count(data) > 0: return
    self._data_list.append(data)
    self.addItem(data.name)

  def clear(self):
    """ Clears the list. """
    super(DataComboBox, self).clear()
    self._data_list.clear()

  def current_data(self):
    """
    :return: a currently selected data file.
    """
    for data in self._data_list:
      if data.name == self.currentText():
        return data

  def __len__(self):
    return len(self._data_list)

  def __getitem__(self, item: int):
    return self._data_list[item]

  def __setitem__(self, key, value):
    raise NotImplementedError("Setting item is not available")


class DataDecisionTable(QTableWidget):
  """ This class represents a table GUI to decide data files corresponding to serials and "SheetData"s

      Attributes:
          _serial_groups: list of "SerialGroup"s
          _sheet_list: list of worksheet names.
          _keyword: a keyword that the name of all worksheet for data has.
  """

  def __init__(self):
    super(DataDecisionTable, self).__init__()
    self._serial_groups: List[SerialGroup] = []
    self._sheet_list: List[str] = []
    self._keyword = ""
    self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
    self.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

  # Getters
  @property
  def serial_groups(self):
    return self._serial_groups

  @property
  def sheet_list(self):
    return self._sheet_list

  @property
  def keyword(self):
    return self._keyword

  # Setters
  @pyqtSlot(list)
  def set_serial_groups(self, tfs: List[TextFile]):
    self._serial_groups = group_data_files(tfs)
    self.update_table()

  @pyqtSlot(ExcelFile)
  def set_sheet_list(self, excel_file: ExcelFile):
    self._sheet_list = []
    with excel_file as excel:
      for sheet in excel.sheets:
        self._sheet_list.append(sheet.name)
    self.update_table()

  @pyqtSlot(str)
  def set_keyword(self, word: str):
    self._keyword = word
    self.update_table()

  def clear(self):
    """ Clears all data. """
    super(DataDecisionTable, self).clear()
    self.setRowCount(0)
    self.setColumnCount(0)

  def update_table(self):
    """ Updates a table when attributes are set.
    """
    if len(self._serial_groups) == 0 or len(self.sheet_list) == 0:
      self.clear()
      return

    self.clear()

    # headers
    vertical_h: List[str] = []
    for sg in self.serial_groups:
      vertical_h.append(sg.serial)

    self.setRowCount(len(self.serial_groups))
    self.setColumnCount(len(self.sheet_list))
    self.setVerticalHeaderLabels(vertical_h)
    self.setHorizontalHeaderLabels(self.sheet_list)

    # cells
    for r, c in product(range(self.rowCount()), range(self.columnCount())):
      cell = DataComboBox()
      if self.keyword == "" or (self.keyword != "" and self.sheet_list[c].count(self.keyword) > 0):
        for tf in self.serial_groups[r]:
          if tf.name.count(self.sheet_list[c].replace(self.keyword, "")) > 0:
            cell.add(tf)
      self.setCellWidget(r, c, cell)

  def get_table(self) -> DataTable:
    """
    :return: a table that contains text files currently selected.
    """
    if len(self._serial_groups) == 0 or len(self.sheet_list) == 0:
      raise Exception("Data table is not updated.")

    result: DataTable = DataTable()

    serials = []
    for sg in self.serial_groups:
      serials.append(sg.serial)
    result.append_header_vs(serials)
    result.append_header_hs(self.sheet_list)

    for r, c in product(range(self.rowCount()), range(self.columnCount())):
      cell: DataComboBox = self.cellWidget(r, c)
      if len(cell) == 0:
        result.insert(None, r, c)
      else:
        result.insert(cell.current_data(), r, c)

    return result


class WgtDataTable(QFrame):
  """ This class represents a widget for a data table.

      Attributes;
          _data_table: "DataDecisionTable" object
          _bt_set_save_file_names: a button that shows a dialog to set the names of saved Excel files.
  """

  def __init__(self):
    super(WgtDataTable, self).__init__()
    self._data_table = DataDecisionTable()
    self._bt_set_save_file_names = QPushButton("Set Save File Names")

    button_layout = QHBoxLayout()
    button_layout.addWidget(self._bt_set_save_file_names)
    button_layout.addSpacerItem(QSpacerItem(70, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))

    layout = QVBoxLayout()
    layout.addWidget(self._data_table)
    layout.addLayout(button_layout)

    self.setLayout(layout)

    self.setFrameShape(QFrame.StyledPanel)

    self._bt_set_save_file_names.clicked.connect(lambda: print(self._data_table.get_table()))

  # Getters
  @property
  def data_table(self):
    return self._data_table

  @property
  def bt_set_save_file_names(self):
    return self._bt_set_save_file_names

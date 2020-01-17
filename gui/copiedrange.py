from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class WgtCopiedRange(QFrame):
  range_changed = pyqtSignal(str)

  def __init__(self):
    super(WgtCopiedRange, self).__init__()
    self._edt_range = QLineEdit()
    self.setStatusTip(
      "If you want to merge data of excel files which will be created from the text files, please type the range of the data.")
    self.setFrameShape(QFrame.StyledPanel)

    layout = QVBoxLayout()
    layout.addWidget(QLabel("Range of merging data (Format: [Sheet name]![From]:[To]"))
    layout.addWidget(self._edt_range)
    self.setLayout(layout)

    self._edt_range.textChanged.connect(lambda: self.range_changed.emit(self._edt_range.text()))

  def excel_range(self):
    return self._edt_range.text()

"""
    This module has a class for GUI to decide a keyword to distinguish data sheets from non-data sheets.
"""

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class WgtKeyword(QFrame):
    """ This class represents a widget to decide a keyword to distinguish data sheets from non-data sheets.

        Attributes:
            _edt_keyboard: QLineEdit object to type a keyword

        Class Attributes:
            keyword_changed: a signal emitted when a keyword is changed.
    """
    keyword_changed = pyqtSignal(str)

    def __init__(self):
        super(WgtKeyword, self).__init__()
        self._edt_keyword = QLineEdit()
        self.setStatusTip("If there is any keyword that sheets for data have, please type here.")
        self.setFrameShape(QFrame.StyledPanel)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Data Sheet Keyword"))
        layout.addWidget(self._edt_keyword)
        self.setLayout(layout)

        self._edt_keyword.setText("_raw")
        self._edt_keyword.textChanged.connect(lambda : self.keyword_changed.emit(self._edt_keyword.text()))

    def keyword(self):
        return self._edt_keyword.text()
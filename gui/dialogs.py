"""
    This module has classes for dialogs.
"""

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from gui.datatable import DataTable
from gui.messages import ErrorMessage, WaitingMessage, InformMessage
from typing import List
from basic.file import text_to_excel
from basic.file.files import TextFile, ExcelFile
from basic.sheetdata.sheetdata import SheetData
from basic.list2d import Table
from itertools import product

class DlgSaveFileName(QDialog):
    """ This class represents a dialog to set the names of saved Excel files.

        Attributes:
            _edt_all_name: QLineEdit object to enter the same name for all saved Excel files.
            _names: a list of QLineEdit objects for each saved Excel file.
            _excel_file: an Excel file.
            _data_table: a table from DataDecisionTable object.
            _sheet_data_list: a list of "SheetData"s
            _bt_start_load: a button to start loading
            _bt_cancel: a button to quit the dialog.
    """
    def __init__(self, excel: ExcelFile, data_table: DataTable, sheet_data_list: List[SheetData]):
        super(DlgSaveFileName, self).__init__()
        self._edt_all_name = QLineEdit()
        self._names: List[QLineEdit] = []
        self._excel_file = excel
        self._data_table = data_table
        self._sheet_data_list = sheet_data_list
        self._bt_start_load = QPushButton("Start Load")
        self._bt_cancel = QPushButton("Cancel")

        name_layout = QFormLayout()
        name_layout.addRow(QLabel('Name of All Files'), self._edt_all_name)
        for serial in self._data_table.header_v:
            edt_name = self.__make_connected_edt()
            self._names.append(edt_name)
            name_layout.addRow(QLabel(serial), edt_name)

        button_layout = QHBoxLayout()
        button_layout.addSpacerItem(QSpacerItem(0, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        button_layout.addWidget(self._bt_start_load)
        button_layout.addWidget(self._bt_cancel)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Excel Name Format: [Serial] [Name] [Date(YYMMDD-HHMM)].xlsx"))
        layout.addLayout(name_layout, 1)
        layout.addLayout(button_layout)

        self.setLayout(layout)

        self._bt_start_load.clicked.connect(self.__bt_start_load_clicked)
        self._bt_cancel.clicked.connect(self.reject)

        self.setWindowModality(Qt.ApplicationModal)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

    def __bt_start_load_clicked(self):
        names = []
        for n in self._names:
            if n.text().count('\\') > 0 \
                    or n.text().count('/') > 0 \
                    or n.text().count(':') > 0 \
                    or n.text().count('*') > 0 \
                    or n.text().count('?') > 0 \
                    or n.text().count('<') > 0 \
                    or n.text().count('>') > 0 \
                    or n.text().count('|') > 0:
                error_mb = ErrorMessage("<nobr>File name must not contain / \\ : * ? < > |</nobr>")
                error_mb.show()
                error_mb.exec_()
                return
            else:
                names.append(n.text())

        loading_mb = WaitingMessage("<nobr>Loading data to excel files...</nobr>")
        loading_mb.show()

        table: Table[TextFile, str, SheetData] = Table[TextFile, str, SheetData]()

        table.append_header_vs(self._data_table.header_v)
        table.append_header_hs(self._sheet_data_list)
        for r, c in product(range(table.n_row), range(table.n_col)):
            table.insert(self._data_table.get(r,c), r, c)

        try:
            text_to_excel(table, self._excel_file, names)
        except Exception as e:
            error_mb = ErrorMessage("<nobr>Error: " + str(e) + "</nobr>")
            error_mb.show()
            error_mb.exec_()
        loading_mb.accept()
        self.accept()

    def __make_connected_edt(self) -> QLineEdit:
        edt_name = QLineEdit()
        self._edt_all_name.textChanged.connect(lambda: edt_name.setText(self._edt_all_name.text()))
        return edt_name


class DlgAboutProgram(QDialog):
    """ This class represents a dialog for information about this program. """
    def __init__(self):
        super(DlgAboutProgram, self).__init__()
        self.resize(1032, 286)
        self.setWindowTitle("About Program")

        self.setWindowModality(Qt.ApplicationModal)

        verticalLayout = QVBoxLayout()
        layout = QFormLayout()
        layout.setHorizontalSpacing(50)
        LBL_PRGM_NAME = QLabel("Program Name")
        LBL_PRGM_NAME.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(0, QFormLayout.LabelRole, LBL_PRGM_NAME)
        PRGM_NAME = QLabel("Text to Excel")
        PRGM_NAME.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(0, QFormLayout.FieldRole, PRGM_NAME)
        LBL_VERSION = QLabel("Version")
        LBL_VERSION.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(1, QFormLayout.LabelRole, LBL_VERSION)
        VERSION = QLabel("1.0.1")
        VERSION.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(1, QFormLayout.FieldRole, VERSION)
        LBL_PRGMMR = QLabel("Programmer")
        LBL_PRGMMR.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(2, QFormLayout.LabelRole, LBL_PRGMMR)
        PRGMMR = QLabel("Younghoon Jeong")
        PRGMMR.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(2, QFormLayout.FieldRole, PRGMMR)
        LBL_E_MAIL = QLabel("E-Mail")
        LBL_E_MAIL.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(3, QFormLayout.LabelRole, LBL_E_MAIL)
        E_MAIL = QLabel("yhjeong0422@gmail.com")
        E_MAIL.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        layout.setWidget(3, QFormLayout.FieldRole, E_MAIL)
        verticalLayout.addLayout(layout)
        LBL_MSG = QLabel("If there is any bug, please send it to my e-mail!")
        LBL_MSG.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
        verticalLayout.addWidget(LBL_MSG)

        self.setLayout(verticalLayout)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
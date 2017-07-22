"""
    This module has classes for GUI to show a list of data (text files).
"""

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from basic.file.files import TextFile
from typing import List
from gui.messages import ErrorMessage
import gui, sys

class DataList(QListWidget):
    """ This class shows a list of data.

        Attributes:
            _data_list: a list of text files

        Class Attributes:
            list_changed: a signal emitted when the list is changed.
    """
    list_changed = pyqtSignal(list)

    def __init__(self):
        super(DataList, self).__init__()
        self._data_list: List[TextFile] = []
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setAcceptDrops(True)
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)

        self.list_changed.connect(self.__print_list)

    # Getters
    @property
    def data_list(self):
        return self._data_list

    def __add_data(self, data: TextFile):
        """
        Adds a data file in the list.
        :param data: a data file.
        """
        if self._data_list.count(data) > 0: return

        self.insertItem(len(self._data_list) - 1, str(data))
        self._data_list.append(data)

    def add_data_list(self, data_list: List[str]):
        """
        Adds data files in the list.
        :param data_list: data files.
        """
        error_files = []

        for data in data_list:
            try:
                self.__add_data(TextFile(data))
            except Exception as e:
                error_files.append((data, str(e)))

        # error
        if len(error_files) != 0:
            error_msg = "<p><nobr>These files are not added for following reason.</nobr><br><br>"
            for ef in error_files:
                error_msg += "<nobr>" + ef[0] + "</nobr><br><nobr>" + "==> " + ef[1] + "</nobr><br><br>"
            error_msg += "</p>"
            try:
                error_mb = ErrorMessage("<nobr>Some files are not added.</nobr>")
                error_mb.setTextFormat(Qt.RichText)
                error_mb.setInformativeText(error_msg)
                error_mb.show()
                error_mb.exec_()
            except Exception as e:
                print(e)

        self.list_changed.emit(self.data_list)


    def remove_selected(self):
        """ Removes selected data files. """
        for item in self.selectedItems():
            self.takeItem(self.row(item))

            for i, data in enumerate(self._data_list):
                if str(data) == item.text():
                    del self._data_list[i]
                    break
        self.list_changed.emit(self.data_list)

    # drag - drop functions for dragging & dropping files from OS Explorer.
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
            self.add_data_list(file_names)
        else:
            event.ignore()

    @pyqtSlot(list)
    def __print_list(self, data_list: List[TextFile]):
        print("data file: " + str(data_list))


class WgtDataList(QFrame):
    """ This class represents a widget for data list GUI.

        Attributes:
            _data_list: "DataList" object
            _bt_add: button to add data files
            _bt_remove: button to remove data files
    """
    def __init__(self):
        super(WgtDataList, self).__init__()
        self._data_list = DataList()
        self._bt_add = QPushButton('Add Data')
        self._bt_remove = QPushButton('Remove Data')

        button_layout = QHBoxLayout()
        button_layout.addWidget(self._bt_add)
        button_layout.addWidget(self._bt_remove)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Data List"))
        layout.addWidget(self._data_list)
        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.setFrameShape(QFrame.StyledPanel)

        self._bt_add.clicked.connect(self.__bt_add_clicked)
        self._bt_remove.clicked.connect(self.__bt_remove_clicked)

    # Getters
    @property
    def data_list(self):
        return self._data_list

    def __bt_add_clicked(self):
        fname, _ = QFileDialog.getOpenFileNames(self, 'Open file', directory = gui.current_directory,
                                                filter="Text File (*.txt)")

        if len(fname) == 0: return
        gui.current_directory = fname[0][:fname[0].rfind('\\')]

        self._data_list.add_data_list(fname)

    def __bt_remove_clicked(self):
        self._data_list.remove_selected()




if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WgtDataList()
    window.show()
    app.exec_()
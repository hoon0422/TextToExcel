""" Classes for files such as normal file, text file, excel file (.xlsx) """

from typing import List
from basic.errors import *
from shutil import copyfile
import os

import xlwings as xw


class File(object):
    """Classes for normal file.

    Attributes:
        _name: file's name without a serial
        _serial: serial of a file. The number of letters of a serial is 12.
        _file_format: the format of a file. It starts with dot(.).
        _path: the directory of a file.
        _full_name: the full name of a file, combination of path, serial, name, and file format.

    Class Variable:
        N_SERIALS: The number of letters of a serial.
    """
    N_SERIALS = 12

    # Constructors
    def __init__(self, full_name: str, using_serial: bool = False):
        """Constructor with full name string.

        :param str full_name: the full name of a file.
        :raise FileNotFoundError: if there is no file matching with "full_name".
        :raise InvalidFileNameError: if "full_name" is invalid file name.
        :raise TooShortFileNameError: if the number of letters of file name
                                      without serials is smaller than "N_SERIALS".
        """
        self._name = ""
        self._serial = ""
        self._using_serial = using_serial
        self._file_format = ""
        self._path = ""
        self._full_name = ""


        # check there is a file.
        try:
            file = open(full_name)
        except:
            raise FileNotFoundError


        # find slash(/, \) symbol for name
        slashIdx = full_name.rfind('/')
        if slashIdx == -1: slashIdx = full_name.rfind('\\')
        if slashIdx == -1: raise InvalidFileNameError("Cannot find '/' symbol.")
        full_name = full_name.replace('/','\\')

        # find dot(.) symbol for format
        dotIdx = full_name.rfind('.')
        if dotIdx == -1: raise InvalidFileNameError("Cannot find '.' symbol.")

        # the number of letters for name must be bigger than the number of serial letters
        if self._using_serial and len(full_name[slashIdx + 1 : dotIdx]) <= File.N_SERIALS:
                raise TooShortFileNameError("File name is too short. Must be bigger than 12")

        # assignment
        self._full_name = full_name
        self._path = full_name[:slashIdx]
        self._file_format = full_name[dotIdx:]
        if using_serial:
            self._name = full_name[slashIdx + 1 + File.N_SERIALS + 1 : dotIdx]
            self._serial = full_name[slashIdx + 1:][:File.N_SERIALS]
        else:
            self._name = full_name[slashIdx + 1: dotIdx]

    # Getters
    @property
    def name(self) -> str:
        return self._name

    @property
    def serial(self) -> str:
        return self._serial

    @property
    def using_serial(self):
        return self._using_serial

    @property
    def file_format(self) -> str:
        return self._file_format

    @property
    def path(self) -> str:
        return self._path

    @property
    def full_name(self) -> str:
        return self._full_name

    # override function
    def __str__(self) -> str:
        return self._serial + ' ' + self._name + self.file_format

    def __eq__(self, o: object) -> bool:
        if isinstance(o, File):
            return self.full_name == o.full_name
        return super().__eq__(o)

    def get_data(self) -> str:
        with open(self._full_name) as f: return f.read()




class TextFile(File):
    """Class for text file(.txt)

    Class variable:
        FORMAT: the format of text file, which is ".txt".
    """
    FORMAT = ".txt"

    def __init__(self, full_name: str, using_serial: bool = True):
        """Constructor with full name string.

        :param str full_name: the full name of a file.
        :raise InvalidFileFormatError: if file format is not ".txt".
        """
        super(TextFile, self).__init__(full_name, using_serial)
        if self._file_format != TextFile.FORMAT:
            raise InvalidFileFormatError("Format must be '.txt', but value is " + self._file_format)


class ExcelFile(File):
    """Class for text file(.xlsx)

    Attributes
        save_file_name: file name that is needed when saving Excel files with new name.
        _current_book: xlwings "Book" object if this object is opened. Otherwise, it is None.

    Class variable:
        FORMAT: the format of Excel file, which is ".xlsx".
    """
    FORMAT = ".xlsx"
    excel_app = None

    @classmethod
    def open_excel_app(cls):
        cls.excel_app = xw.App()
        cls.excel_app.visible = False
        cls.excel_app.display_alerts = False
        print("Excel " + str(cls.excel_app) + " is opened.")

    @classmethod
    def close_excel_app(cls):
        for book in cls.excel_app.books:
            book.close()
        cls.excel_app.quit()
        print("Excel " + str(cls.excel_app) + " is closed.")
        cls.excel_app = None

    def __init__(self, full_name: str):
        """Constructor with full name string.

        :param str full_name: the full name of a file.
        :raise InvalidFileFormatError: if file format is not ".xlsx".
        """
        super(ExcelFile, self).__init__(full_name, False)
        if self._file_format != ExcelFile.FORMAT:
            raise InvalidFileFormatError("Format must be '.xlsx', but value is " + self._file_format)
        self.save_file_name = ""
        self._current_book = None
        self._temp_excel = None

    def open(self) -> xw.Book:
        """Opens xlwings "Book" object with "_full_name".
        
        :return: opened "Book" object.
        :raise OSError: if another program use the file.
        """
        for app in xw.apps:
            for book in app.books:
                if book.fullname.upper() == self._full_name.upper():
                    raise OSError("Another program use the file. Please close the program.")

        self._temp_excel = self._path + "\\~$" + self._name + "_~$TEMP~$" + self.FORMAT
        try:
            os.remove(self._temp_excel)
        except: pass
        copyfile(self._full_name, self._temp_excel)
        os.popen('attrib +h ' + self._temp_excel).close() # hide temp file
        self._current_book = ExcelFile.excel_app.books.open(self._temp_excel)
        return self._current_book

    def close(self) -> None:
        """ Close xlwings "Book" object. """
        if self._current_book is not None:
            if self.save_file_name == "": # overwrite source Excel file.
                self._current_book.save(self._full_name)
            elif self.save_file_name.find('\\') == -1 \
                    and self.save_file_name.find('/') == -1: # same directory with source file.
                self._current_book.save(self._path + '\\' + self.save_file_name)
            else:
                self._current_book.save(self.save_file_name)
            self._current_book.close()
        self._current_book = None
        os.remove(self._temp_excel)
        self._temp_excel = None

    def __enter__(self):
        """ Returns "open" function's result.

        :return: "open" function's result.
        """
        return self.open()

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Executes "close" function."""
        self.close()


class SerialGroup(object):
    def __init__(self):
        self._serial = ""
        self._data_list: List[TextFile] = []

    @property
    def serial(self):
        return self._serial

    def __len__(self) -> int:
        return len(self._data_list)

    def __getitem__(self, item: int) -> TextFile:
        return self._data_list[item]

    def __setitem__(self, index: int, value: TextFile) -> None:
        self.__checkFile(value)
        self._data_list.__setitem__(index, value)

    def __delitem__(self, key) -> None:
        self._data_list.__delitem__(key)
        if len(self._data_list) == 0:
            self._serial = ""

    def count(self, file: TextFile) -> int:
        i = 0
        for f in self._data_list:
            if file == f: i += 1
        return i
        
    def clear(self) -> None:
        self._data_list.clear()

    def insert(self, index: int, value: TextFile) -> None:
        self.__checkFile(value)
        self._data_list.insert(index, value)

    def append(self, value: TextFile):
        self.__checkFile(value)
        self._data_list.append(value)

    def __checkFile(self, file: TextFile) -> None:
        if isinstance(file, TextFile) is False: raise TypeError(file.__str__() + " is not 'TextFile' object.")
        if file.using_serial is False: raise NoSerialExistError(file.__str__() + " does not use serial.")

        if self._serial == "":
            self._serial = file.serial
        elif self._serial != file.serial:
            raise SerialNotMatchedError(file.__str__() + "'s serial (" + file.serial
                                        + ") is not same with group's serial")





if __name__ == "__main__":
    pass
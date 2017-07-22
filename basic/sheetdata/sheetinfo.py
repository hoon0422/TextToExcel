"""
    This module has classes for information when manipulating a worksheet.
"""

import abc
from xlwings import Sheet
from typing import TypeVar, Generic, List
from itertools import product

S = TypeVar('S')
class SheetInfo(Generic[S]):
    """ This abstract class represents information of how to manipulate a worksheet.

        Attribute:
            _infom_name: the name of an information.
    """
    __metaclass__ = abc.ABCMeta

    def __init__(self, info_name: str):
        self._info_name = info_name

    # Getters
    @property
    def info_name(self):
        return self._info_name

    @property
    @abc.abstractmethod
    def info_type(self) -> type:
        """
        :return: A type of information data. For example, if data type is str, return str.
        """
        pass

    @abc.abstractmethod
    def apply_info_to_sheet(self, sheet: Sheet, value: S):
        """
        Manipulates a worksheet with the specified value.
        :param sheet: a worksheet which will be manipulated
        :param value: a value used when manipulating
        """
        pass


class MissingValueInfo(SheetInfo[str]):
    """ This class tells how to manipulate a worksheet in terms of missing values. """
    def __init__(self):
        super(MissingValueInfo, self).__init__("Missing Value Info")

    def apply_info_to_sheet(self, sheet: Sheet, value: str):
        """
        Manipulates a worksheet with the specified value.
        It removes the value in the worksheet.
        :param sheet: a worksheet which will be manipulated
        :param value: a value used when manipulating. In this case, missing values. If there are more than one
                    value, it will separate with white spaces.
        """
        missing_values = value.split()

        r = sheet.range('A1').end('right').end('down')
        rows, cols = r.row, r.column
        if rows == 1048576:
            rows = (0 if sheet.range('A1').value == None else 1)
        if cols == 16384:
            cols = (0 if sheet.range('A1').value == None else 1)

        for row, col in product(range(rows), range(cols)):
            v = sheet.cells[row, col].value
            for mv in missing_values:
                try:
                    if float(mv) == v:
                        sheet.range((row + 1, col + 1)).value = None
                except:
                    if mv == v:
                        sheet.range((row + 1, col + 1)).value = None

    @property
    def info_type(self) -> type:
        """
        :return: A type of information data. For example, if data type is str, return str.
        """
        return str


sheet_infos: List[SheetInfo] = [] # global variable that contains "SheetInfo"s.
sheet_infos.append(MissingValueInfo())
"""
    This module has classes for Excel worksheet.
"""

from basic.sheetdata.sheetinfo import sheet_infos
from basic.sheetdata.sheetinfo import SheetInfo

class SheetData(object):
    """ This class represents an Excel worksheet.

        Attributes:
            _sheet_name: the name of a worksheet
            _sheet_infos: a dictionary whose key is info type and value is the value of info.
    """
    def __init__(self, sheet_name: str):
        self._sheet_name = sheet_name
        self._sheet_infos = {}
        for info in sheet_infos:
            self._sheet_infos[type(info)] = None

    # Getters
    @property
    def sheet_name(self):
        return self._sheet_name

    def __getitem__(self, item: SheetInfo):
        return self._sheet_infos[item]

    def __setitem__(self, key: SheetInfo, value):
        if isinstance(type(key.info_type), type(value)): raise TypeError("Sheet info type is not matched with value.")
        self._sheet_infos[key] = value
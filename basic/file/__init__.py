"""
 This package has functions and classes for files.
"""
import re
from datetime import datetime
from itertools import product
from typing import List

import xlwings as xw
from basic.sheetdata.sheetdata import SheetData

from basic.errors import *
from basic.file.files import TextFile, ExcelFile, SerialGroup
from basic.list2d import Matrix, Table
from basic.sheetdata.sheetinfo import sheet_infos

__all__ = ["files", "group_data_files", "text_to_excel", "merge_specified_range"]


def str_to_matrix(s: str) -> Matrix[str]:
  """
  Makes string data to matrix. Delimiter is white spaces.
  :param s: string data
  :return: matrix that contains data
  """
  result: Matrix[str] = Matrix[str]()

  row = 0
  for line in s.split('\n'):
    col = 0
    line_exist = False

    for cell in line.split('\t'):
      result.insert(cell, row, col)
      line_exist = True
      col += 1

    if line_exist:
      row += 1

  return result


def group_data_files(data_files: List[TextFile]) -> List[SerialGroup]:
  """
  Groups a list of text files by serials.
  :param data_files: a list of text files
  :return: a list of "SerialGroup"s contains the text files which is a parameter.
  """
  serial_group_list: List[SerialGroup] = []
  for tf in data_files:
    inserted = False
    for serial_group in serial_group_list:
      try:
        serial_group.append(tf)
        inserted = True
        break
      except SerialNotMatchedError:
        pass

    if inserted is False:
      serial_group_list.append(SerialGroup())
      serial_group_list[len(serial_group_list) - 1].append(tf)

  return serial_group_list


def make_data_table(serial_groups: List[SerialGroup], sheet_data: List[SheetData], data_sheet_keyword: str = "") \
    -> Table[List[TextFile], str, SheetData]:
  """
  Returns a table from a list of "SerialGroup"s and a list of "SheetData"s using keyword.
  :param serial_groups: a list of "SerialGroup"s
  :param sheet_data: a list of "SheetData"s
  :param data_sheet_keyword: keyword that only data sheets have.
  :return: a table that contains a list of text files in one cell corresponding to "SerialGroup" and "SheetData"
  """
  table: Table[List[TextFile], str, SheetData] = Table[List[TextFile], str, SheetData]()

  serials: List[str] = []
  for sg in serial_groups:
    serials.append(sg.serial)

  table.append_header_vs(serials)
  table.append_header_hs(sheet_data)
  for r, c in product(range(len(serials)), range(len(sheet_data))):
    table.insert([], r, c)

  for sg in serial_groups:
    for tf, sd in product(sg, sheet_data):
      sd_name = sd.sheet_name
      if data_sheet_keyword == "" or sd_name.count(data_sheet_keyword) > 0:
        sd_name = sd_name.replace(data_sheet_keyword, "").strip()
        if tf.name.count(sd_name) > 0:
          table.get_with_header(sg.serial, sd).append(tf)

  return table


def text_to_excel(data_table: Table[TextFile, str, SheetData], excel_file: ExcelFile, save_names: List[str]):
  """
  Load data in text files to an Excel file.
  :param data_table: a table that contains text files, with a vertical header consisting of serials, and with
                      a horizontal header consisting of "SheetData"s
  :param excel_file: Excel file
  :param save_names: a list of file names for new Excel files.
  """
  path = excel_file.path
  final_names = []

  for serial, save_name in zip(data_table.header_v, save_names):
    with excel_file as excel:
      for sd in data_table.header_h:
        tf = data_table.get_with_header(serial, sd)
        sheet = excel.sheets[sd.sheet_name]
        if tf is not None:
          sheet.range((1, 1)).value = str_to_matrix(tf.get_data()).contents()
          sheet.autofit('r')
          path = tf.path
        for si in sheet_infos:
          si.apply_info_to_sheet(sheet, sd[si])
      file_name = path + "\\" + serial + " " + save_name \
                  + (" " if len(save_name) != 0 else "") \
                  + datetime.now().strftime("%y%m%d-%H%M") + ".xlsx"
      final_names.append(file_name)
      excel_file.save_file_name = file_name

  return final_names


def merge_specified_range(excel_file_names: List[str], excel_range: str, save_name: str, path=''):
  p = re.compile(
    r"([^:\\/?*\[\]]{1,31}\'!|[^:\\/?*\[\]]{1,31}!)(\$?[a-z]{1,3}\$?[0-9]{1,7}(:\$?[a-z]{1,3}\$?[0-9]{1,7})?|\$[a-z]{1,3}:\$[a-z]{1,3}|[a-z]{1,3}:[a-z]{1,3}|\$[0-9]{1,7}:\$[0-9]{1,7}|[0-9]{1,7}:[0-9]{1,7}|[a-z_\\][a-z0-9_.]{0,254})")
  if not p.match(excel_range):
    raise Exception("'" + excel_range + "' is not a valid Excel range address.")

  merged = xw.Book()
  sheet_name = excel_range[0:excel_range.find('!')]

  current_row = 1
  for excel_file_name in excel_file_names:
    excel_file = xw.books.open(path + excel_file_name)
    copied_range = excel_file.sheets[sheet_name].range(excel_range)
    merged.sheets[0].range('A' + str(current_row)).value = copied_range.value
    current_row += copied_range.rows.count
  print(path + save_name)
  merged.save(path + save_name)

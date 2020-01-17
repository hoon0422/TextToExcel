"""
    This module has exception classes.
"""


class InvalidFileNameError(Exception):
  """ This exception happens when a file name is invalid. """

  def __init__(self, msg: str):
    super(InvalidFileNameError, self).__init__()
    self._msg = msg

  def __str__(self):
    return self._msg


class InvalidFileFormatError(Exception):
  """ This exception happens when a file format is invalid. """

  def __init__(self, msg: str):
    super(InvalidFileFormatError, self).__init__()
    self._msg = msg

  def __str__(self):
    return self._msg


class TooShortFileNameError(Exception):
  """ This exception happens when a file name is too short. """

  def __init__(self, msg: str):
    super(TooShortFileNameError, self).__init__()
    self._msg = msg

  def __str__(self):
    return self._msg


class SerialNotMatchedError(Exception):
  """
      This exception happens when a file is added in a "SerialGroup"
      which has a different serial from the file's.
  """

  def __init__(self, msg: str):
    super(SerialNotMatchedError, self).__init__()
    self._msg = msg

  def __str__(self):
    return self._msg


class NoSerialExistError(Exception):
  """ This exception happens when a file that does not use serial is added in a "SerialGroup". """

  def __init__(self, msg: str):
    super(NoSerialExistError, self).__init__()
    self._msg = msg

  def __str__(self):
    return self._msg

"""
    This module has classes for message boxes.
"""

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class WaitingMessage(QMessageBox):
    """ This message box shows a message that something is on the process. """
    def __init__(self, msg: str):
        super(WaitingMessage, self).__init__()
        self.setWindowModality(Qt.ApplicationModal)
        self.setStandardButtons(QMessageBox.NoButton)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setWindowTitle("Loading")
        self.setText(msg)

class ErrorMessage(QMessageBox):
    """ This message box shows a message that something is not processed due to some errors. """
    def __init__(self, msg: str):
        super(ErrorMessage, self).__init__()
        self.setWindowModality(Qt.ApplicationModal)
        self.setStandardButtons(QMessageBox.Ok)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setIcon(QMessageBox.Critical)
        self.setWindowTitle("Error")
        self.setText(msg)
        self.setTextInteractionFlags(Qt.TextSelectableByMouse)

class InformMessage(QMessageBox):
    """ This message box shows a message to inform something. """
    def __init__(self, msg: str):
        super(InformMessage, self).__init__()
        self.setWindowModality(Qt.ApplicationModal)
        self.setStandardButtons(QMessageBox.Ok)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setIcon(QMessageBox.Information)
        self.setWindowTitle("Inform")
        self.setText(msg)
        self.setTextInteractionFlags(Qt.TextSelectableByMouse)

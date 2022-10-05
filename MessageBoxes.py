from multiprocessing.connection import wait
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
import re
from PyQt5 import *
from PyQt5 import QtCore
from PyQt5 import QtGui

def MessageBoxErreur():
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Critical)
    msgBox.setStandardButtons(QMessageBox.Ok)
    msgBox.setWindowFlag(Qt.FramelessWindowHint)
    msgBox.setMinimumWidth(500)
    msgBox.setFixedWidth(600)
    msgBox.setWindowOpacity(0.85)
    msgBox.setStyleSheet("QPushButton{border-radius:10px;border:2px solid rgb(99, 126, 193);background-color: rgb(85, 96, 124);font: 87 14pt ""Segoe UI"";width:60px;color: rgb(255, 255, 255);}QPushButton:hover{background-color: rgb(114, 153, 146);}QLabel{font: 87 14pt ""Segoe UI"";color: white;}QMessageBox{border-radius:25px;border:7px solid red;background-color: rgb(32, 33, 53);}")
    return msgBox

def MessageBoxValid():
    msgBox = QMessageBox()
    msgBox.setIcon(QMessageBox.Information)
    msgBox.setStandardButtons(QMessageBox.Ok)
    msgBox.setWindowFlag(Qt.FramelessWindowHint)
    msgBox.setMinimumWidth(500)
    msgBox.setFixedWidth(600)
    msgBox.setWindowOpacity(0.85)
    msgBox.setStyleSheet("QPushButton{border-radius:10px;border:2px solid rgb(99, 126, 193);background-color: rgb(85, 96, 124);font: 87 14pt ""Segoe UI"";width:60px;color: rgb(255, 255, 255);}QPushButton:hover{background-color: rgb(114, 153, 146);}QLabel{font: 87 14pt ""Segoe UI"";color: white;}QMessageBox{border-radius:25px;border:7px solid rgb(95,158,160);background-color: rgb(32, 33, 53);}")
    return msgBox

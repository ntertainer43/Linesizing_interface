import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from pathlib import Path
import xlwings as xw
import Parse


class ELabel(QLabel):
    """docstring for Elabel."""

    def __init__(self,parent,i,j):
        super(ELabel, self).__init__(parent)

        self.streamnum = i
        self.rownum = j

class ELineEdit(QLineEdit):
    """docstring for Elabel."""

    def __init__(self,parent,i,j):
        super(ELineEdit, self).__init__(parent)

        self.streamnum = i
        self.rownum = j


class PipeWindow(QMainWindow):
    """docstring for PipeWindow."""

    def __init__(self, parent, id):
        super(PipeWindow, self).__init__(parent)

        self.widget =  QWidget()
        self.layout =  QGridLayout(self.widget)
        self.book = Parse.Sizing()

        self.setCentralWidget(self.widget)



        #elements in the PipeWindow
        self.widget_format()

        self.setWindowTitle(str(id))
        self.show()

    def widget_format(self):
        for i in range (5):
            label_property = ELabel(self.book.prop[i],1,i )
            label_unit = ELabel(self.book.unit[i],1,i)
            text_physical = ELineEdit(str(self.book.propvalue[0][i+1]),1,i)
            self.layout.addWidget(label_property)
            self.layout.addWidget(text_physical)
            self.layout.addWidget(label_unit)
            text_physical.editingFinished.connect(self.callExcel)

    def callExcel(self):
        new_value = float(self.sender().text())
        i = self.sender().rownum
        j = self.sender().streamnum
        #print(float(self.sender().text())+20)
        self.book.writestream( new_value,i,j)

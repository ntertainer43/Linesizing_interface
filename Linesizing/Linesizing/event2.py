import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from pathlib import Path
import xlwings as xw


class ELabel(QLabel):
    """docstring for Elabel."""

    def __init__(self,parent,i,j):
        super(ELabel, self).__init__(parent)

        self.streamnum = i
        self.tagnum = j

class ELineEdit(QLineEdit):
    """docstring for Elabel."""

    def __init__(self,parent,i,j):
        super(ELineEdit, self).__init__(parent)

        self.streamnum = i
        self.tagnum = j


class PipeWindow(QMainWindow):
    """docstring for PipeWindow."""

    def __init__(self):
        super(PipeWindow, self).__init__()

        self.widget =  QWidget()
        self.layout =  QGridLayout(self.widget)
        self.book = Sizing('n')

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
            self.layout.addWidget(label_unit)
            self.layout.addWidget(text_physical)



# the class for carrying out excel calculation and interfacing
class Sizing(object):
   """docstring for ."""

   def __init__(self):
       super().__init__()


class Sizing(object):
    """docstring for Sizing."""

    def __init__(self, flpath):
        super(Sizing, self).__init__()
        self.flpath = 'C:/Users/capta/Desktop/Linesizing/Linesizing/HMB.xlsx'
        self.wb = xw.Book(self.flpath)
        self.sh = self.wb.sheets['Property Calc_Min. Flow']
        self.streamlist = []
        self.prop = self.sh.range('C4:C13').value
        self.unit = self.sh.range('D4:D13').value
        print(self.unit)
        self.propvalue =[]
        self.parsestream()

    def parsestream(self):
        j = 5
        while  self.sh.range(2,j).value != None:
            self.readstream(j)
            self.streamlist.append(self.sh.range(2,j).value)
            j += 6


    def readstream(self, j):
        a =[]
        for i in range (3,13):
            a.append(self.sh.range(i,j).value)
            #print(a)
        self.propvalue.append(a)
        #print(self.propvalue)

    def writestream(self,changed_value,i=0,j=5):
        excel_row = j+3
        excel_column = i*6 +5
        self.sh.range(excel_row, excel_column).value= changed_value



def main():
    app =  QApplication(sys.argv)
    wind = PipeWindow()


    app.exec_()

def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

if __name__ == "__main__":

    import sys
    sys.exit(main())
    sys.excepthook = except_hook

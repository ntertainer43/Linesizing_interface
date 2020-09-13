import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from pathlib import Path
import xlwings as xw

fname =''

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

class MainWindow(QMainWindow):

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.pipe_count = 0
# instanciate widget from other class
# just have this code always makes it easy still dont know how it woks
#        self.form_widget = FormWidget(self)
#        _widget =  QWidget()
#        _layout =  QVBoxLayout(_widget)
#        _layout.addWidget(self.form_widget)
#        self.setCentralWidget(_widget)

        self.widget =  QWidget()
        self.layout =  QGridLayout(self.widget)
        self.frame = QFrame()
        self.setCentralWidget(self.widget)


        #Instating sub tools uing QActions variables  for the open dialogue box

        openFile = QAction('Open', self)
        openFile.setShortcut('Ctrl+O')
        openFile.setStatusTip('Open new File')
        openFile.triggered.connect(self.showDialog)

        pipe = QAction('Pipe', self)
        pipe.setStatusTip('Create a new pipe segment')
        pipe.triggered.connect(self.pipeCreate)


        # menubar creation
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(openFile)

        # toolbar creation
        toolbar = QToolBar('Mymaintoolbar')
        self.addToolBar(toolbar)
        toolbar.addAction(pipe)




        self.statusbar = self.statusBar()
        self.setAcceptDrops(True)
        self.setGeometry(300, 300, 550, 450)
        self.show()


        #the function to show the dialgoue box for browse
    def showDialog(self):
        home_dir = str(Path.home())
        fname = QFileDialog.getOpenFileName(self, 'Open file', home_dir)
        print(fname[0])
        self.xl = Sizing(fname[0])

    def pipeCreate(self):
        self.pipe_count += 1
        print (self.pipe_count)
        self.btn1 = Pipes("Pipe 1",self, self.pipe_count)
        print(self.btn1)
        self.layout.addWidget(self.btn1)

#drag and drop feature not yet implemented
    def dragEnterEvent(self, e):
        e.accept()

    def dropEvent(self, e):
        position = e.pos()
        self.btn1.move(position)

        e.setDropAction(Qt.MoveAction)
        e.accept()


class Pipes(QPushButton):

    def __init__(self, title, parent, id):
        super().__init__(title, parent)

        self.pipe_id = id
# function for drag and drop feature currently not working
    def mouseMoveEvent(self, e):

        if e.buttons() != Qt.RightButton:
            return

        mimeData = QMimeData()

        drag = QDrag(self)
        drag.setMimeData(mimeData)
        drag.setHotSpot(e.pos() - self.rect().topLeft())

        dropAction = drag.exec_(Qt.MoveAction)

    def mousePressEvent(self, e):

        super().mousePressEvent(e)

        if e.button() == Qt.LeftButton:
            pipwind = PipeWindow(self, self.pipe_id)




class PipeWindow(QMainWindow):
    """docstring for PipeWindow."""

    def __init__(self, parent, id):
        super(PipeWindow, self).__init__(parent)

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
    win = MainWindow()
    p = Sizing('nimesh')
    p.writestream(2,5,0)
    app.exec_()

def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

if __name__ == "__main__":

    import sys
    sys.exit(main())
    sys.excepthook = except_hook

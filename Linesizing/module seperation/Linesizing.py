import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QIcon
from pathlib import Path
#import xlwings as xw
import Parse
from Dialogue import PipeWindow


class MainWindow(QMainWindow):

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.pipe_count = 0
# instanciate widget from other class
# just have this code always makes it easy still dont know how it woks
#        self.form_widget = FormWidget(self)
#        _widget =  QWidget()0
#        _layout =  QVBoxLayout(_widget)
#        _layout.addWidget(self.form_widget)
#        self.setCentralWidget(_widget)2

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



def main():
    app =  QApplication(sys.argv)
    win = MainWindow()
    app.exec_()

def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)

if __name__ == "__main__":

    import sys
    sys.exit(main())
    sys.excepthook = except_hook

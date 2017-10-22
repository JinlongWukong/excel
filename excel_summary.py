# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled_main.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!
############### Revison history ####################
# verison :1.0  built on 2017.10.22

from PyQt4 import QtCore, QtGui
from PyQt4.QtGui import *
import workflow
import winsound

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_MainWindow(QtGui.QMainWindow):
    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.setupUi(self)
        self.retranslateUi(self)

    def setupUi(self, MainWindow):
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(1120, 838)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.plainTextEdit = QtGui.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setGeometry(QtCore.QRect(20, 60, 1081, 761))
        self.plainTextEdit.setObjectName(_fromUtf8("plainTextEdit"))
        self.verticalScrollBar = QtGui.QScrollBar(self.centralwidget)
        self.verticalScrollBar.setStyleSheet(_fromUtf8("    background: #8000FF;\n"
                                                       "    border: 3px solid grey;\n"
                                                       "    border-radius:5px;\n"
                                                       "    min-height: 20px;"))
        self.plainTextEdit.setVerticalScrollBar(self.verticalScrollBar)
        self.lineEdit = QtGui.QLineEdit(MainWindow)
        self.lineEdit.setGeometry(QtCore.QRect(810, 25, 281, 27))
        self.lineEdit.setObjectName(_fromUtf8("lineEdit"))
        self.label = QtGui.QLabel(MainWindow)
        self.label.setGeometry(QtCore.QRect(640, 25, 121, 21))
        self.label.setObjectName(_fromUtf8("label"))
        self.label_version = QtGui.QLabel(MainWindow)
        self.label_version.setGeometry(QtCore.QRect(850, 820, 250, 15))
        self.label_version.setObjectName(_fromUtf8("label"))
        #self.pushButton = QtGui.QPushButton(self.centralwidget)
        #self.pushButton.setGeometry(QtCore.QRect(850, 20, 112, 34))
        #self.pushButton.setObjectName(_fromUtf8("pushButton"))
        #self.pushButton.clicked.connect(self.browse)
        #self.comboBox = QtGui.QComboBox(self.centralwidget)
        #self.comboBox.setGeometry(QtCore.QRect(20, 20, 761, 27))
        #self.comboBox.setObjectName(_fromUtf8("comboBox"))
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtGui.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1120, 31))
        self.menubar.setObjectName(_fromUtf8("menubar"))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtGui.QStatusBar(MainWindow)
        self.statusbar.setObjectName(_fromUtf8("statusbar"))
        MainWindow.setStatusBar(self.statusbar)

        #redirect console to plainText
        redir = redirect(self.plainTextEdit)
        sys.stdout = redir
        sys.stderr = redir

        # Enable dragging and dropping onto the GUI
        self.setAcceptDrops(True)
        # Enable Copy&Paste
        self.setupEditActions()
        self.actionCut.setEnabled(False)
        self.actionCopy.setEnabled(False)
        self.actionCut.triggered.connect(self.plainTextEdit.cut)
        self.actionCopy.triggered.connect(self.plainTextEdit.copy)
        self.actionPaste.triggered.connect(self.plainTextEdit.paste)
        self.plainTextEdit.copyAvailable.connect(self.actionCut.setEnabled)
        self.plainTextEdit.copyAvailable.connect(self.actionCopy.setEnabled)
        QtGui.QApplication.clipboard().dataChanged.connect(self.clipboardDataChanged)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(_translate("MainWindow", "Excel_Tool", None))
        #self.pushButton.setText(_translate("MainWindow", "Brower", None))
        self.lineEdit.setText(_translate("Dialog", "统计值.xlsx", None))
        self.label.setText(_translate("Dialog", "TargetFileName：", None))
        self.label_version.setText(_translate("Dialog", "Version: 1.0, built on 20171022", None))

    def browse(self):
        #fileName = QtGui.QFileDialog.getOpenFileName(self, "Open File", '.',"HTML Files (*.htm *.html)")
        fileName = QtGui.QFileDialog.getOpenFileName(self, "Open File", '.', "EXCEL Files (*.xlsx)")
        if fileName:
            if self.comboBox.findText(fileName) == -1:
                self.comboBox.addItem(fileName)
            self.comboBox.setCurrentIndex(self.comboBox.findText(fileName))
            self.start()

    def start(self):
        if self.comboBox.currentText() == '':
            print("Please brower a valid html")
        else:
            self.plainTextEdit.moveCursor(QtGui.QTextCursor.End)
            self.th = workThread(self.comboBox.currentText())
            self.th.started.connect(self.th.worker)
            self.th.start()

    # The following three methods set up dragging and dropping for the app
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
            self.plainTextEdit.moveCursor(QtGui.QTextCursor.End)
            for url in e.mimeData().urls():
                fname = unicode(url.toLocalFile())
                print("Input file: " + fname)
                tfile = unicode(self.lineEdit.text())
                self.th = workThread(fname,tfile)
                self.th.started.connect(self.th.worker)
                self.th.start()
                break
        else:
            e.ignore()

    def setupEditActions(self):
        self.actionCopy = QtGui.QAction(
                "&Copy", self, priority=QtGui.QAction.LowPriority,
                shortcut=QtGui.QKeySequence.Copy)

        self.actionCut = QtGui.QAction(
                "&Cut", self, priority=QtGui.QAction.LowPriority,
                shortcut=QtGui.QKeySequence.Cut)

        self.actionPaste = QtGui.QAction(
                "&Paste", self, priority=QtGui.QAction.LowPriority,
                shortcut=QtGui.QKeySequence.Paste,
                enabled=(len(QtGui.QApplication.clipboard().text()) != 0))

    def clipboardDataChanged(self):
        self.actionPaste.setEnabled(len(QtGui.QApplication.clipboard().text()) != 0)

class redirect(object):
    def __init__(self, text):
        self.output = text

    def write(self, string, color = False):
        fmt = QTextCharFormat()
        if color:
            fmt.setForeground(QColor("red"))
        else:
            fmt.setForeground(QColor("black"))

        self.output.mergeCurrentCharFormat(fmt)
        self.output.insertPlainText(string)
        self.output.ensureCursorVisible()

class workThread(QtCore.QThread):
    def __init__(self, file = None, target = 'summary.xlsx'):
        super(workThread, self).__init__()
        self.file = file
        self.targetFile = target

    def worker(self):
        winsound.PlaySound('alert', winsound.SND_ASYNC)
        workflow.worker(self.file, self.targetFile)
        self.exit()

if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    splash = QSplashScreen(QPixmap("spinner.gif"))
    splash.show()
    splash.showMessage("loading...")
    app.processEvents() #in case GUI not response, enable processEvent
    ui = Ui_MainWindow()
    ui.show()
    splash.finish(ui)
    sys.exit(app.exec_())


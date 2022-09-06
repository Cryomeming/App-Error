# The Imported items.
import os
import sys

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ';' + os.environ['PATH']

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, \
                            QVBoxLayout, QLabel
from PyQt5.QtCore import Qt

from subprocess import call


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)


class button(QPushButton):  # Draw a Button
    def __init__(self, label_text):
        super().__init__()
        self.setText(label_text)
        self.setStyleSheet('''
            font-size: 17.5px;
            width: 30px;
            height: 30px;
        ''')
        # Please don't touch.

class AppUI(QWidget):  # Make UI
    def __init__(self):
        super().__init__()
        self.resize(500, 500)
        self.initUI()

    def initUI(self):
        mainLayout = QVBoxLayout()
        self.setWindowFlag(Qt.FramelessWindowHint)
        self.setLayout(mainLayout)

        self.placeholder = QLabel("WORD, EXCEL, PDF Converter and Merger")
        self.placeholder.move(70, -200)
        mainLayout.addChildWidget(self.placeholder)

        self.COPYRIGHT = QLabel(" Copyright © Houjie Li 2022, Version: 1.0")
        self.COPYRIGHT.move(70, 225)
        mainLayout.addChildWidget(self.COPYRIGHT)

        self.buttonTest1 = button("WORD to PDF")
        self.buttonTest1.resize(300, 70)
        self.buttonTest1.move(100, 70)
        self.buttonTest1.clicked.connect(self.WTOPAppCaller)  # Call WORDTOPDF Function
        mainLayout.addChildWidget(self.buttonTest1)

        self.buttonTest2 = button("EXCEL to PDF")
        self.buttonTest2.resize(300, 70)
        self.buttonTest2.move(100, 170)
        self.buttonTest2.clicked.connect(self.ETOPAppCaller)  # Call EXCELTOPDF Function
        mainLayout.addChildWidget(self.buttonTest2)

        self.buttonTest3 = button("MERGE PDF")
        self.buttonTest3.resize(300, 70)
        self.buttonTest3.move(100, 270)
        self.buttonTest3.clicked.connect(self.PTOPAppCaller)  # Call PDFMERGER function
        mainLayout.addChildWidget(self.buttonTest3)

        self.buttonExit = button('EXIT')
        self.buttonExit.resize(300, 70)
        self.buttonExit.move(100, 370)

        self.buttonExit.clicked.connect(self.exit)
        mainLayout.addChildWidget(self.buttonExit)

    def WTOPAppCaller(self):
        call(["python", "WORDTOPDF.py"])

    def ETOPAppCaller(self):
        call(["python", "EXCELTOPDF.py"])

    def PTOPAppCaller(self):
        call(["python", "MERGEPDF.py"])

    def exit(self):
        QApplication.quit()


 ## How the Desktop Application is able to show itself. ##
if __name__ == '__main__':  # Please don't touch this man ok? #
    app = QApplication(sys.argv)
    # app.setStyle("fusion")
    app.setStyleSheet('''
        QWidget {
            font-size: 20px;
        }
    ''')

    myNewApp = AppUI()
    myNewApp.show()

    sys.exit(app.exec_())


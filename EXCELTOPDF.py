# The Imported items.
import os
import sys
from datetime import datetime

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ';' + os.environ['PATH']

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QListWidget, \
    QVBoxLayout, QMessageBox, QAbstractItemView
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter

from aspose.cells import Workbook, SaveFormat  # Excel Converter`

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)

class ListWidget(QListWidget):
    # Queue Placeholder Text #
    def __init__(self):
        super().__init__()
        self._placeholder_text = ''

    @property
    def placeholder_text(self):
        return self._placeholder_text

    @placeholder_text.setter
    def placeholder_text(self, text):
        self._placeholder_text = text
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.count() == 0:
            painter = QPainter(self.viewport())
            painter.save()

            color = self.palette().placeholderText().color()
            painter.setPen(color)

            font_metrics = self.fontMetrics()
            elided_text = font_metrics.elidedText(
                self.placeholder_text,
                Qt.ElideRight,
                self.viewport().width()
            )
            painter.drawText(self.viewport().rect(), Qt.AlignCenter, elided_text)
            painter.restore()

    # How Drag-and-Drop works for the Queue #
    # IMPORTANT: Please don't touch. It'll break the code! #

    def __init__(self, parent=None):
        super().__init__(parent=None)
        self.setAcceptDrops(True)
        self.setStyleSheet('''font-size:25px''')
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    def dialogMessage(self, message, windowtitle, detailedtext, icon, button, escapebutton):  # Conversion Dialog Message
        dialog = QMessageBox(self)
        dialog.setWindowTitle(windowtitle)
        dialog.setText(message)
        dialog.setStandardButtons(button)
        dialog.setDetailedText(detailedtext)
        dialog.setIcon(icon)
        dialog.setEscapeButton(escapebutton)
        dialog.show()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            return super().dragEnterEvent(event)
            # Returns to original event state in dragEnterEvent class

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            return super().dragMoveEvent(event)
            # Returns to original event state in dragMoveEvent

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            userFiles1 = []
            # What the Queue is: Empty (at the moment)
            # Files are dropped into the queue(userFiles1)

            for url in event.mimeData().urls():
                if url.isLocalFile() and url.toString():
                    userFiles1.append(str(url.toLocalFile()))  # Add file to list

                    if url.toLocalFile().endswith(".xlsx"):  #IF item in list is xlsx file
                        workbookconverter = Workbook(str(url.toLocalFile()))

                        localtime = datetime.now()

                        current_time = localtime.strftime("%m%d%H%M%S")

                        workbookconverter.save("../Converter/output/" + "XTOP_" + str(current_time) + ".pdf", SaveFormat.PDF)

                        self.dialogMessage(message="PDFs are in 'output' folder.", windowtitle='Excel files sucessfully converted.',
                                           detailedtext="The name of your PDF will follow this format: 'XTOP_MMDDHHMMSS'",
                                           button=QMessageBox.Ok, escapebutton=QMessageBox.Ok,
                                           icon=QMessageBox.Information)

            self.addItems(userFiles1)
        else:
            return super().dropEvent(event) # Returns to the original event state in dropEvent class


class button(QPushButton):
    def __init__(self, label_text):
        super().__init__()
        self.setText(label_text)
        self.setStyleSheet('''
            font-size: 17.5px;
            width: 170px;
            height: 30;
        ''')
        # Please don't touch.

class AppDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Excel to PDF Converter.')
        # self.setWindowIcon(QIcon(resource_path(r'PDFFile.ico')))
        self.resize(500, 500)
        self.initUI()

    def initUI(self):
        mainLayout = QVBoxLayout()

        ### mainLayout Format ###

        # Layout Format: Placeholder Text #
        self.fileQueue = ListWidget()
        mainLayout.addWidget(self.fileQueue)
        self.fileQueue.placeholder_text = 'Drop your Excel files here.'

        self.setLayout(mainLayout)

 ## How the Desktop Application is able to show itself. ##
if __name__ == '__main__':  # Please don't touch this man ok? #
    app = QApplication(sys.argv)
    # app.setStyle("fusion")
    app.setStyleSheet('''
        QWidget {
            font-size: 20px;
        }
    ''')

    myNewApp = AppDemo()
    myNewApp.show()

    sys.exit(app.exec_())
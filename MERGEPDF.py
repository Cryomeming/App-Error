# The Imported items.
import os
import sys

if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ';' + os.environ['PATH']

from PyQt5.QtWidgets import QApplication, QWidget, QLineEdit, QPushButton, QListWidget, \
    QVBoxLayout, QHBoxLayout, QComboBox, QStylePainter, \
    QFileDialog, QMessageBox, QAbstractItemView, QStyle, QStyleOptionComboBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPainter, QPalette

from PyPDF2 import PdfFileMerger  # PDF Merger

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

            self.addItems(userFiles1)  # Keep the self.addItems where the for url is positioned for items to not be cloned when hovering over two or more.

        else:
            return super().dropEvent(event) # Returns to the original event state in dropEvent class


class output_field(QLineEdit):
    def __init__(self):
        super().__init__()
        # self.setMinimumHeight(50)
        self.height = 55
        self.setStyleSheet('''font-size: 20px;''')
        # It's recommended you keep this the same.

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()
            # This section cannot use the return!
            # We already made a line dedicated to return.
            # Using ignore should be easier

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()
            # Same Reason as Above Comment

    def dropEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()

            if event.mimeData().urls():
                self.setText(event.mimeData().urls()[0].toLocalFile())
                # This line of code demonstrates...
                # How the file is able to tell you the location of the file in your Desktop.
        else:
            event.ignore()
            # Reason stays the same as previous comment.

class button(QPushButton):
    def __init__(self, label_text):
        super().__init__()
        self.setText(label_text)
        self.setStyleSheet('''
            font-size: 17.5px;
            width: 150px;
            height: 30px;
        ''')
        # Please don't touch.

class AppDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PDF Merger.')
        self.resize(500, 500)
        self.initUI()

    def initUI(self):
        outputFolderRow = QHBoxLayout()
        mainLayout = QVBoxLayout()
        buttonLayout = QHBoxLayout()

        self.outputFile = output_field()
        outputFolderRow.addWidget(self.outputFile)

        ### The buttonLayout Buttons ###

        # buttonLayout: File Explorer #
        self.buttonFileExplorer = button('SAVE Location')
        self.buttonFileExplorer.clicked.connect(self.FileExplorer)
        buttonLayout.addWidget(self.buttonFileExplorer, Qt.AlignLeft)

        # buttonLayout: Merge #
        self.buttonMerge = button('Merge PDFs')
        self.buttonMerge.clicked.connect(self.PDFMerger)
        buttonLayout.addWidget(self.buttonMerge)

        # buttonLayout: Clearing Queue #
        self.buttonEXIT = button('EXIT')
        # self.buttonReset.clicked.connect(self.clearQueue)
        self.buttonEXIT.clicked.connect(self.EXITapp)
        buttonLayout.addWidget(self.buttonEXIT)


        ### mainLayout Format ###

        # Layout Format: Placeholder Text #
        self.fileQueue = ListWidget()
        mainLayout.addWidget(self.fileQueue)
        self.fileQueue.placeholder_text = 'Drop your PDF files here.'

        # Layout Format: Part 2 #
        mainLayout.addLayout(outputFolderRow)
        mainLayout.addLayout(buttonLayout)

        self.setLayout(mainLayout)

    ## Button Functions ##
    def clearQueue(self):
        self.outputFile.setText('')
        self.fileQueue.clear()

    def EXITapp(self):
        QApplication.quit()

    def populateFilename(self):
        path = self._getSaveFilePath()
        if path:
            self.outputFile.setText(path)

    # Connects to dialogTesting Button
    def dialogMessage(self, message, windowtitle, detailedtext, icon, button, escapebutton):
        dialog = QMessageBox(self)
        dialog.setWindowTitle(windowtitle)
        dialog.setText(message)
        dialog.setStandardButtons(button)
        dialog.setDetailedText(detailedtext)
        dialog.setIcon(icon)
        dialog.setEscapeButton(escapebutton)
        dialog.show()

    # Main Functions
    def _getSaveFilePath(self):
        file_save_path, _ = QFileDialog.getSaveFileName(self, 'Save file', os.getcwd(), "PDF File (*.pdf*)")
        return file_save_path

    def FileExplorer(self):
        if not self.outputFile.text():
            # self.dialogMessage("Empty File Directory")
            self.populateFilename()
            return

    def PDFMerger(self):
        if self.fileQueue.count() > 0:
            pdfMerger = PdfFileMerger()

            try:
                for i in range(self.fileQueue.count()):
                    pdfMerger.append(self.fileQueue.item(i).text())

                pdfMerger.write(self.outputFile.text())
                pdfMerger.close()

                self.fileQueue.clear()
                self.dialogMessage(message="Merge Complete.",
                                   windowtitle="Completed Merging.",
                                   detailedtext="Please look for your PDF in the output folder.",
                                   icon=QMessageBox.Information,
                                   button=QMessageBox.Ok,
                                   escapebutton=QMessageBox.Ok)

            except Exception:
                self.dialogMessage(message="Merge Unsuccessful.", windowtitle="Application Error",
                                   detailedtext="An item was missing and/or the PDF file you have placed in is considered a zero-byte file. This is an issue with the PDFs apparently, do not know why.",
                                   icon=QMessageBox.Critical,
                                   button=QMessageBox.Ok,
                                   escapebutton=QMessageBox.Ok)

            # use this sparingly. This was only used for testing purposes. #
            # except Exception as e:
            #     self.dialogMessage(e,e,e,e,e)

        else:
            self.dialogMessage(message="Nothing was sensed by the program.",
                               windowtitle="No files detected.", detailedtext="The queue is empty, put some PDFs in.",
                               icon=QMessageBox.Warning,
                               button=QMessageBox.Cancel,
                               escapebutton=QMessageBox.Cancel)


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
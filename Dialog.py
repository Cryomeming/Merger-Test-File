# NOTE: This is for pure testing purposes.

import sys, os, io
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ';' + os.environ['PATH']
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QListWidget, \
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QStylePainter, \
    QDialog, QFileDialog, QMessageBox, QAbstractItemView, QStyle, QStyleOptionComboBox
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QIcon, QPainter, QPalette
from PyPDF2 import PdfFileMerger
import aspose.words as aw


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)

# Placeholder Text for QComboBox
class FileChoiceDropList(QComboBox):
    def __init__(self):
        super().__init__()
        self.changed = False

    def paintEvent(self, event):

        painter = QStylePainter(self)
        painter.setPen(self.palette().color(QPalette.Text))

        # draw the combobox frame, focusrect and selected etc.
        opt = QStyleOptionComboBox()
        self.initStyleOption(opt)
        painter.drawComplexControl(QStyle.CC_ComboBox, opt)

        if self.currentIndex() < 0:
            opt.palette.setBrush(
                QPalette.ButtonText,
                opt.palette.brush(QPalette.ButtonText).color().lighter(),
            )
            if self.placeholderText():
                opt.currentText = self.placeholderText()

        # draw the icon and text
        painter.drawControl(QStyle.CE_ComboBoxLabel, opt)

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

            for url in event.mimeData().urls():
                if url.isLocalFile():
                    if url.toString():
                        userFiles1.append(str(url.toLocalFile()))
            self.addItems(userFiles1)
            # How Files are able to be dropped into the Queue

        else:
            return super().dropEvent(event)
            # Returns to the original event state in dropEvent class

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
            width: 170px;
            height: 30;
        ''')
        # basically a button format.

class AppDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Desktop Merge Testing Platform App')
        self.setWindowIcon(QIcon(resource_path(r'PDF.ico')))
        self.resize(1850, 850)
        self.initUI()

    def initUI(self):
        mainLayout = QVBoxLayout()
        outputFolderRow = QHBoxLayout()
        buttonLayout = QHBoxLayout()

        self.outputFile = output_field()
        outputFolderRow.addWidget(self.outputFile)

        ### The OutputFolderRow Buttons ###

        # OutputFolderRow: PDF Merge Button #
        self.buttonMerge = button('Merge your PDFs')
        self.buttonMerge.clicked.connect(self.mergePDFTest)
        outputFolderRow.addWidget(self.buttonMerge)

        self.buttonMerge2 = button('DOCX Test merge')
        outputFolderRow.addWidget(self.buttonMerge2)

        # buttonLayout: Drop-down List #
        self.dropdownlist = FileChoiceDropList()
        self.dropdownlist.addItems(['.docx', '.pdf'])
        self.dropdownlist.setPlaceholderText('Choose File Format')
        self.dropdownlist.setCurrentIndex(-1)

        buttonLayout.addWidget(self.dropdownlist)

        ### mainLayout Format ###

        # Layout Format: Part 1 #
        mainLayout.addLayout(outputFolderRow)

        # Layout Format: Placeholder Text #
        self.fileQueue = ListWidget()
        mainLayout.addWidget(self.fileQueue)
        self.fileQueue.placeholder_text = 'Drag Files here.'

        mainLayout.addLayout(buttonLayout)

        self.setLayout(mainLayout)

    ## Button Functions ##
    # Connects to dialogTesting Button
    def dialogMessage(self, message, detailedtext, icon, button, escapebutton):
        dialog = QMessageBox(self)
        dialog.setWindowTitle('Sample Text')
        dialog.setText(message)
        dialog.setStandardButtons(button)
        dialog.setDetailedText(detailedtext)
        dialog.setIcon(icon)
        dialog.setEscapeButton(escapebutton)
        dialog.show()

    def mergePDFTest(self):
        if self.fileQueue.count() > 0:
            pdfMerger = PdfFileMerger()

            try:
                for i in range(self.fileQueue.count()):
                    pdfMerger.append(self.fileQueue.item(i).text())

                pdfMerger.write(self.outputFile.text())
                pdfMerger.close()

                self.fileQueue.clear()
                self.dialogMessage("Completed thing", "Please thank the people of PyPDF2 for creating this incredible module", icon=QMessageBox.Information, button=QMessageBox.Ok, escapebutton=QMessageBox.Ok)

            except Exception:
                self.dialogMessage("PDFs were unable to merge properly.", "Something was missing within the program. Please add more to fix the issue.", icon=QMessageBox.Critical, button=QMessageBox.Ok, escapebutton=QMessageBox.Ok)

            # use this sparingly. This was only used for testing purposes. #
            #except Exception as e:
                #self.dialogMessage(e,e,e,e,e)

        else:
            self.dialogMessage("Nothing was sensed by the program.", "During the code, something wasn't set properly. Either test the program using the debugger to find out the problem.", icon=QMessageBox.Warning, button=QMessageBox.Cancel, escapebutton=QMessageBox.Cancel)


    def mergeDOCXTest(self):
        if self.fileQueue.count() > 0:
            testing = aw.Document()
            testing.remove_all_children()




    # TODO: How does mergeFile get called?
        ## When mergeButton is pressed.
    # TODO: How does mergeFile know what userFiles is
        ## We can make a List as Arguement, and run it that way
    # TODO: When is mergeFile called?
        ## When mergeButton is pressed.


 ## How the Desktop Application is able to show itself. ##
if __name__ == '__main__':
    # Please don't touch this man ok? #
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

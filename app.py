import traceback
import sys

from PyQt5 import QtWidgets
from PyQt5.QtCore import QObject, pyqtSignal, QThread
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from openpyxl.utils.exceptions import InvalidFileException

from ui import Ui_MainWindow
from validator import EmailValidator


class App(QtWidgets.QMainWindow):

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.init_ui()

        self.ui.pushButton.clicked.connect(self.load_file)

    def load_file(self):
        dialog = QFileDialog.getOpenFileName(self, "Select Excel file to import", "", "Excel (*.xls *.xlsx)")
        file_name = dialog[0]

        self.start_validate(file_name)

    @staticmethod
    def popup_info():
        msg = QMessageBox()
        msg.setWindowTitle('Информация')
        msg.setText('Для начала работы выберите файл!')

        x = msg.exec_()

    def show_end(self):
        self.ui.label_2.setText('Файл записан')
        self.ui.pushButton.setEnabled(1)

    def init_ui(self):
        self.setWindowTitle('EmailValidator')

    def start_validate(self, file_name):
        self.obj = Validator(file_name)
        self.t = QThread()
        self.obj.moveToThread(self.t)
        self.t.started.connect(self.obj.start)
        self.obj.finishSignal.connect(self.t.quit)
        self.obj.finishSignal.connect(self.show_end)

        self.t.start()

        self.ui.pushButton.setEnabled(0)


class Validator(QObject):
    finishSignal = pyqtSignal()

    def __init__(self, file_name):
        super().__init__()
        self.file_name = file_name

    def start(self):
        try:
            email_validator = EmailValidator(self.file_name)

            addresses = email_validator.get_info()
            check_result = email_validator.check_address(addresses)
            email_validator.write_data(check_result)

            self.finishSignal.emit()

        except InvalidFileException:
            App.popup_info()


def main():

    def log_uncaught_exceptions(ex_cls, ex, tb):
        text = '{}: {}:\n'.format(ex_cls.__name__, ex)
        text += ''.join(traceback.format_tb(tb))

        print(text)
        QtWidgets.QMessageBox.critical(None, 'Error', text)
        quit()

    sys.excepthook = log_uncaught_exceptions

    app = QtWidgets.QApplication(sys.argv)
    application = App()
    application.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

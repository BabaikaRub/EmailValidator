import traceback
import sys

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox, QFileDialog

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

        if dialog:
            self.ui.label_2.setText('Файл загружен.')

            file_name = dialog[0]

            self.start_validate(file_name)

    def init_ui(self):
        self.setWindowTitle('EmailValidator')

    def popup_end(self):
        msg = QMessageBox()
        msg.setWindowTitle('Информация')
        msg.setText('Файл успешно записан!!!')

        x = msg.exec_()

    def start_validate(self, file_name):
        email_validator = EmailValidator(file_name)

        addresses = email_validator.get_info()
        check_result = email_validator.check_address(addresses)
        flag = email_validator.write_data(check_result)

        if flag:
            self.popup_end()


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

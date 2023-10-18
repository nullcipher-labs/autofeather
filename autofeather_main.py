from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication
import sys
import autofeather


class FadedApp(QtWidgets.QMainWindow, autofeather.Ui_MainWindow):
    """a class for the gui window of our app, PyQt5 requirement"""

    def __init__(self, parent=None):
        super(FadedApp, self).__init__(parent)
        self.setupUi(self)
        self.setWindowTitle('Automatic Faded Edges')


def main():
    app = QApplication(sys.argv)
    form = FadedApp()
    form.show()
    app.exec_()


if __name__ == '__main__':
    main()

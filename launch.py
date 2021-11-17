from PyQt6.QtWidgets import *
from grk_parser.widgets import MainWindow


if __name__ == '__main__':
    app = QApplication(["GRK Data Synchronizer"])
    window = MainWindow()
    window.show()
    app.exec()

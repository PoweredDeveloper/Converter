import sys
from PySide6 import QtWidgets

from modules.window import Window

if __name__ == '__main__':
    app = QtWidgets.QApplication([])

    widget = Window()
    widget.resize(1200, 800)
    widget.show()

    sys.exit(app.exec())
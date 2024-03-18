import sys
from MainScreen import MainScreen
from PyQt6.QtWidgets import QApplication


from Logger import logger as log


class Main:
    def __init__(self):
        #super().__init__()
        self.class_name = "Main"
        app = QApplication(sys.argv)
        window = MainScreen()
        window.show()
        sys.exit(app.exec())


if __name__ == "__main__":
    main = Main()

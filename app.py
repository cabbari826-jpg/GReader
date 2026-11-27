import sys
from PySide6.QtWidgets import QApplication
from ordertool.gui import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("PDF Order Converter & Validator")
    app.setOrganizationName("Local Business Tools")
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

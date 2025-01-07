import sys
from PyQt5.QtWidgets import QApplication
from app.app import PDFViewerApp

def main():
    app = QApplication(sys.argv)
    viewer = PDFViewerApp()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()

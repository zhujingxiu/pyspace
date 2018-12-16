import sys
import uiLinqForm
from PyQt5.QtWidgets import QApplication, QMainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = uiLinqForm.Ui_LinqForm()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

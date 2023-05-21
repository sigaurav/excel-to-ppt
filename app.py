import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QFileDialog, QListWidget, QPushButton
from PyQt5 import QtCore
from main import Generate_PPT


class ExcelBrowser(QMainWindow):

    file_path = ""
    selected_sheets_names = []

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Browser")
        self.setGeometry(100, 100, 400, 300)

        self.file_path = None
        self.sheet_names = []

        self.central_widget = QWidget()
        self.main_layout = QVBoxLayout()
        self.central_widget.setLayout(self.main_layout)

        self.file_button = QPushButton("Browse")
        self.file_button.clicked.connect(self.browse_file)

        self.generate_button = QPushButton("Generate PPT")
        self.generate_button.clicked.connect(self.generate_ppt)

        self.sheet_list = QListWidget()
        self.sheet_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.sheet_list.itemClicked.connect(self.toggle_sheet_selection)


        self.main_layout.addWidget(self.file_button)
        self.main_layout.addWidget(self.sheet_list)
        self.main_layout.addWidget(self.generate_button)

        self.setCentralWidget(self.central_widget)

    def browse_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        self.file_path = file_path

        if file_path:
            self.file_path = file_path
            self.load_sheets()

    def load_sheets(self):
        import pandas as pd

        self.sheet_list.clear()

        try:
            xls = pd.ExcelFile(self.file_path)
            self.sheet_names = xls.sheet_names

            self.sheet_list.addItems(self.sheet_names)

            for i in range(self.sheet_list.count()):
                item = self.sheet_list.item(i)
                item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
                item.setCheckState(QtCore.Qt.Checked)

        except Exception as e:
            print(f"Error loading sheets: {e}")

    def toggle_sheet_selection(self, item):
        current_state = item.checkState()

        if current_state == QtCore.Qt.Checked:
            item.setCheckState(QtCore.Qt.Unchecked)
        else:
            item.setCheckState(QtCore.Qt.Checked)

    def selected_sheets(self):
        selected_sheets = []
        for i in range(self.sheet_list.count()):
            item = self.sheet_list.item(i)

            if item.checkState() == QtCore.Qt.Checked:
                selected_sheets.append(item.text())

        self.selected_sheets_names = selected_sheets

        return selected_sheets

    def generate_ppt(self):
        # Call the generate_file() function from the "main" file here
        file_path = self.file_path
        selected_sheets_names = self.selected_sheets()
        Generate_PPT().generate_ppt(file_path, selected_sheets_names)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelBrowser()
    window.show()
    sys.exit(app.exec_())

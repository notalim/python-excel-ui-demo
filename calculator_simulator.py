import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QGridLayout, QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QVBoxLayout, QWidget
from PyQt5.QtCore import Qt
import os

# This is needed to mute the terminal output on macOS (didn't work half of the time?? \_(ツ)_/¯)
os.environ["QT_MAC_WANTS_LAYER"] = "1"

# main class
class CalculatorSimulator(QMainWindow):
    # constructor
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Calculator Simulator")
        self.setGeometry(100, 100, 375, 200)

        self.filename = ""

        self.init_ui()

    # initialize the UI
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QGridLayout(central_widget)

        self.row1_label = QLabel("Row 1:", self)
        self.row1_edit = QLineEdit(self)
        self.row1_edit.setFixedWidth(50)

        layout.addWidget(self.row1_label, 0, 0)
        layout.addWidget(self.row1_edit, 0, 1)

        self.row2_label = QLabel("Row 2:", self)
        self.row2_edit = QLineEdit(self)
        self.row2_edit.setFixedWidth(50)
        layout.addWidget(self.row2_label, 1, 0)
        layout.addWidget(self.row2_edit, 1, 1)

        self.column1_label = QLabel("Column 1:", self)
        self.column1_edit = QLineEdit(self)
        self.column1_edit.setFixedWidth(50)
        layout.addWidget(self.column1_label, 0, 2)
        layout.addWidget(self.column1_edit, 0, 3)

        self.column2_label = QLabel("Column 2:", self)
        self.column2_edit = QLineEdit(self)
        self.column2_edit.setFixedWidth(50)
        layout.addWidget(self.column2_label, 1, 2)
        layout.addWidget(self.column2_edit, 1, 3)

        self.add_button = QPushButton("Add", self)
        self.add_button.clicked.connect(self.add)
        layout.addWidget(self.add_button, 2, 0)

        self.subtract_button = QPushButton("Subtract", self)
        self.subtract_button.clicked.connect(self.subtract)
        layout.addWidget(self.subtract_button, 2, 1)

        self.result_label = QLabel("Answer:", self)
        self.result_value_label = QLabel("", self)
        layout.addWidget(self.result_label, 3, 0)
        layout.addWidget(self.result_value_label, 3, 1)

        self.load_excel_button = QPushButton("Load Excel", self)
        self.load_excel_button.clicked.connect(self.load_excel)
        layout.addWidget(self.load_excel_button, 4, 0)

        self.generate_excel_button = QPushButton("Generate Excel", self)
        self.generate_excel_button.clicked.connect(self.generate_excel)
        layout.addWidget(self.generate_excel_button, 4, 1)

    # load excel file
    def load_excel(self):
        self.filename, _ = QFileDialog.getOpenFileName(self, "Open Excel file", "", "Excel files (*.xlsx)")
    
    # add
    def add(self):
        self.perform_operation("+")

    # subtract
    def subtract(self):
        self.perform_operation("-")

    # perform operation
    def perform_operation(self, operation):
        if not self.filename:
            QMessageBox.warning(self, "Error", "Please load an Excel file first.")
            return

        try:
            row1 = int(self.row1_edit.text())
            row2 = int(self.row2_edit.text())
            column1 = int(self.column1_edit.text())
            column2 = int(self.column2_edit.text())
        except ValueError:
            QMessageBox.warning(self, "Error", "Please enter valid row and column numbers.")
            return

        df = pd.read_excel(self.filename)
        try:
            value1 = df.iat[row1 - 1, column1 - 1]
            value2 = df.iat[row2 - 1, column2 - 1]
        except IndexError:
            QMessageBox.warning(self, "Error", "Row or column index out of bounds.")
            return

        if operation == "+":
            result = value1 + value2
        elif operation == "-":
            result = value1 - value2

        self.result_value_label.setText(f"{result}")

    # generate excel file
    def generate_excel(self):
        nrows, ncols = 10, 10
        data = np.random.randint(0, 100, size=(nrows, ncols))


        col_titles = [f"Column {i + 1}" for i in range(ncols)]

        df = pd.DataFrame(data, columns=col_titles)

        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        filename, _ = QFileDialog.getSaveFileName(self, "Save Excel file", "", "Excel files (*.xlsx)", options=options)

        if not filename:
            return

        with pd.ExcelWriter(filename) as writer:
            df.to_excel(writer, index=False)

        self.filename = filename


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = CalculatorSimulator()
    window.show()

    sys.exit(app.exec_())

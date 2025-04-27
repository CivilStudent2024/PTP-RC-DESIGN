import sys
import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QPushButton, QFileDialog

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Display Excel Data in QTableWidget")
        self.setGeometry(100, 100, 800, 600)

        # Set up the layout and central widget
        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)

        # Create the QTableWidget
        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        # Create a button to load Excel file
        self.load_button = QPushButton("Load Excel File", self)
        layout.addWidget(self.load_button)

        # Set the central widget
        self.setCentralWidget(central_widget)

        # Connect the button click to the function
        self.load_button.clicked.connect(self.load_excel_data)

    def load_excel_data(self):
        # Open a file dialog to select an Excel file
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_dialog.setNameFilter("Excel Files (*.xls *.xlsx)")
        file_dialog.setViewMode(QFileDialog.ViewMode.List)

        if file_dialog.exec():
            file_path = file_dialog.selectedFiles()[0]
            print(file_path)

            # Read the Excel file using pandas
            df = pd.read_excel(file_path,sheet_name='Beam16')

            # Set the table row and column count based on DataFrame
            self.table_widget.setRowCount(df.shape[0])
            self.table_widget.setColumnCount(df.shape[1])

            # Set the table headers from DataFrame columns
            self.table_widget.setHorizontalHeaderLabels(df.columns)

            # Insert the data into the QTableWidget
            for row in range(df.shape[0]):
                for col in range(df.shape[1]):
                    item = QTableWidgetItem(str(df.iloc[row, col]))
                    self.table_widget.setItem(row, col, item)

# Initialize the PyQt application
app = QApplication(sys.argv)

# Create and show the main window
window = MainWindow()
window.show()

# Start the PyQt event loop
sys.exit(app.exec())

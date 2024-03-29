import pandas as pd
from PyQt6.QtWidgets import QMainWindow, QTextEdit, QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem
from datetime import datetime
import base64
import time
from Logger import logger as log

class MainScreen(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.class_name = "MainScreen"
        self.setFixedSize(1200, 800) # Set the fixed size of the QMainWindow
        self.today_date = datetime.today().strftime('%d-%b-%Y')
        self.today_day, self.today_month, self.today_year = self.today_date.split('-')
        self.initUI()
        self.edited_data = None

    def initUI(self):
        layout = QVBoxLayout()

        self.setLayout(layout)
        self.setWindowTitle('Excel Viewer')

        # Create a central widget
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Create a layout for the central widget
        layout = QVBoxLayout()

        # Create a button to open the file dialog
        self.openFileButton = QPushButton('Open Excel File')
        self.openFileButton.clicked.connect(self.openFile)
        layout.addWidget(self.openFileButton)

        # Create a table widget to display the data
        self.tableWidget = QTableWidget()
        self.tableWidget.cellChanged.connect(self.cellChanged)  # Connect the cellChanged signal
        layout.addWidget(self.tableWidget)



        # Create Save and Close buttons
        #start_button = QPushButton("Start")
        self.downloadButton = QPushButton('Download Edited Data')
        close_button = QPushButton("Close")

        # Connect button click events to functions
        #start_button.clicked.connect(self.start)
        self.downloadButton.clicked.connect(self.downloadData)
        self.downloadButton.setEnabled(False)  # Initially disable the download button
        close_button.clicked.connect(self.close)


        # Add buttons to the layout
        #layout.addWidget(start_button)
        layout.addWidget(self.downloadButton)
        layout.addWidget(close_button)
        # Set the layout for the central widget
        central_widget.setLayout(layout)



    def start(self):
        log.info(f"{self.class_name} Started")

    def openFile(self):

        # Open a file dialog to select the Excel file
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open Excel File', '', 'Excel Files (*.xlsx *.xls)')

        if file_path:
            # Read the Excel file using pandas
            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                log.info(f"{self.class_name} Error reading Excel file: {e}")
                return

            if df is not None:  # Check if data frame is not None
                self.edited_data = df.copy()  # Store a copy of the original data for editing

                # Clear the table widget
                self.tableWidget.setRowCount(0)
                self.tableWidget.setColumnCount(len(df.columns))
                self.tableWidget.setHorizontalHeaderLabels(df.columns)

                # Populate the table widget with data from the DataFrame
                for row in range(len(df)):
                    self.tableWidget.insertRow(row)
                    for col in range(len(df.columns)):
                        item = QTableWidgetItem(str(df.iloc[row, col]))
                        self.tableWidget.setItem(row, col, item)

                self.tableWidget.resizeColumnsToContents()
                self.downloadButton.setEnabled(True)  # Enable the download button
        else:
            log.info(f"{self.class_name} - Failed to read Excel file.")

    def cellChanged(self, row, column):
        # Update the edited_data DataFrame with the new value
        self.edited_data.iloc[row, column] = self.tableWidget.item(row, column).text()

    def downloadData(self):
        # Open a file dialog to save the edited data
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save Edited Data', '', 'Excel Files (*.xlsx)')

        if file_path:
            try:
                self.edited_data.to_excel(file_path, index=False)
                log.info(f"{self.class_name} - Edited data saved successfully: {file_path}")
            except Exception as e:
                log.info(f"{self.class_name} - Error saving edited data: {e}")


    def closeApp(self):
        self.close()
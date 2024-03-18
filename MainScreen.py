import pandas as pd
from PyQt6.QtWidgets import QMainWindow, QTextEdit, QLabel, QApplication, QLineEdit, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem
from PyQt6.QtCore import Qt
from datetime import datetime
import base64
import time
from Logger import logger as log
import pyperclip as ppc

class MainScreen(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.class_name = "MainScreen"
        self.setFixedSize(1200, 800) # Set the fixed size of the QMainWindow
        self.today_date = datetime.today().strftime('%d-%b-%Y')
        self.today_day, self.today_month, self.today_year = self.today_date.split('-')
        self.initUI()
        self.edited_data = None
        self.table_name = ''

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

        # Create a QLabel widget
        self.table_label = QLabel("Table Name", self)

        # Create a QLineEdit widget
        self.table_edit = QLineEdit(self)
        self.table_edit.setPlaceholderText("enter table name")
        
        # Connect a signal to handle table changes
        self.table_edit.textChanged.connect(self.on_table_changed)
        
        # Create a QLabel widget
        self.query_label = QLabel("Generated Query", self)

        # Create a QTextEdit widget
        self.sql_query_edit = QTextEdit()
        # Set the number of visible lines (height) to 5
        self.sql_query_edit.setFixedHeight(self.sql_query_edit.fontMetrics().lineSpacing() * 10)
        # Enable vertical scroll bar
        #self.sql_query_edit.setVerticalScrollBarPolicy(0x1)  # QtCore.Qt.ScrollBarAlwaysOn
        
        # Enable vertical scroll bar
        self.sql_query_edit.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)

        # Force the vertical scrollbar to always show, because when QTextEdit is disabled scrollbar not showing
        self.sql_query_edit.setStyleSheet("QTextEdit {overflow: scroll;}")
        
        self.sql_query_edit.setPlainText("")
        self.sql_query_edit.setDisabled(True)


        # Create Save and Close buttons
        self.frame_query_button = QPushButton("Frame Query")
        self.downloadButton = QPushButton('Download Excel Data')
        self.close_button = QPushButton("Close")
        self.copy_button = QPushButton("Copy Query")

        # Connect button click events to functions
        self.frame_query_button.clicked.connect(self.frameQuery)
        self.downloadButton.clicked.connect(self.downloadData)
        self.downloadButton.setEnabled(False)  # Initially disable the download button
        self.close_button.clicked.connect(self.close)
        self.copy_button.clicked.connect(self.copyQuery)
        self.copy_button.setEnabled(False)  # Initially disable the download button


        # Add buttons to the layout
        layout.addWidget(self.table_label)
        layout.addWidget(self.table_edit)
        layout.addWidget(self.query_label)
        layout.addWidget(self.sql_query_edit)
        layout.addWidget(self.frame_query_button)
        layout.addWidget(self.copy_button)
        layout.addWidget(self.downloadButton)
        layout.addWidget(self.copy_button)
        layout.addWidget(self.close_button)
        # Set the layout for the central widget
        central_widget.setLayout(layout)


    def frameQuery(self):
        log.info(f"{self.class_name} Frame Query")
        self.sql_query_edit.setDisabled(False)
        self.copy_button.setEnabled(True)


        log.info(self.edited_data)
        # Constructing the INSERT query
        query_builder = []
        query_builder.append("INSERT INTO ")
        query_builder.append(self.table_name)
        query_builder.append(" (")

        # Append column names
        query_builder.append(", ".join(self.edited_data.columns))
        query_builder.append(")")
        query_builder.append("VALUES")

        # Append values
        for index, row in self.edited_data.iterrows():
            values = "('" + "', '".join(map(str, row)) + "')"
            query_builder.append(values + ",")

        # Remove the trailing comma
        query_builder[-1] = query_builder[-1][:-1]

        # Combine all lines into a single query
        query = '\n'.join(query_builder) + ';'

        log.info('Final query: '+query)
        self.sql_query_edit.setText(query)

        #is_enabled = not self.sql_query_edit.isEnabled()
        #self.sql_query_edit.setEnabled(is_enabled)
        #self.toggle_button.setText("Finaly Query" if not is_enabled else "Please upload excel")

    def copyQuery(self):
        ppc.copy(self.sql_query_edit.toPlainText())
        self.copy_button.setEnabled(False)
    
    def on_table_changed(self, tableName):
        log.info('Table name:'+tableName)
        self.table_name = tableName


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


# 1 2 3 4 5 6 7

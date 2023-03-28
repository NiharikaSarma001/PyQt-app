import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

# Initialize Firebase app with credentials
cred = credentials.Certificate("D:\pandastask\pandastask-cac90-firebase-adminsdk-wb8e4-e3b772c5fc.json")
firebase_admin.initialize_app(cred)

# Get a Firestore client object
db = firestore.client()

import pandas as pd

# Read Excel file into pandas DataFrame
df = pd.read_excel("C:/Users/Admin/Downloads/sample_data.xlsx", sheet_name='Sheet1')

# #iterate over each row
# for index, row in df.iterrows():
#     doc_ref = db.collection('my_collection').document()
#     doc_ref.set(row.to_dict())

#concatenate the existing data frame
#new_df = pd.concat([df]*500, ignore_index=True)

# Convert DataFrame to dictionary
data = df.to_dict(orient='records')

import threading

# Upload data to Firestore using threading
def upload_to_firestore(documents):
    for doc in documents:
        db.collection('student_info').add(doc)

thread = threading.Thread(target=upload_to_firestore, args=(data,))
thread.start()

from PyQt5 import QtCore
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QMainWindow, QLabel, QCheckBox, QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QDialog, QTableWidget, QTableWidgetItem
class ExcelImporter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Importer")
        self.setGeometry(100, 100, 500, 300)
        self.label = QLabel("Click the button to import Excel file.", self)
        self.label.move(100, 50)
        self.button = QPushButton("Import Excel", self)
        self.button.move(100, 100)
        self.button.clicked.connect(self.import_excel)
        self.show()

    def import_excel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)", options=options)
        if file_name:
            try:
            # Read Excel file
                df = pd.read_excel(file_name, sheet_name="Sheet1", header=0)

            # Store file name as instance variable
                self.file_name = file_name

            # Create checkboxes for each column
                checkboxes_layout = QVBoxLayout()
                self.checkboxes = []
                for col in df.columns:
                    checkbox = QCheckBox(col)
                    self.checkboxes.append(checkbox)
                    checkboxes_layout.addWidget(checkbox)
                checkboxes_widget = QWidget()
                checkboxes_widget.setLayout(checkboxes_layout)

            # Create OK and Cancel buttons
                ok_button = QPushButton("OK")
                ok_button.clicked.connect(self.select_columns)
                cancel_button = QPushButton("Cancel")
                cancel_button.clicked.connect(self.close)

            # Add checkboxes and buttons to a layout
                buttons_layout = QHBoxLayout()
                buttons_layout.addWidget(ok_button)
                buttons_layout.addWidget(cancel_button)
                main_layout = QVBoxLayout()
                main_layout.addWidget(self.label)
                main_layout.addWidget(checkboxes_widget)
                main_layout.addLayout(buttons_layout)

            # Create a dialog and show it
                dialog = QDialog(self)
                dialog.setLayout(main_layout)
                dialog.setWindowTitle("Select Columns to Import")
                dialog.setGeometry(200, 200, 300, 300)
                dialog.exec_()
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))
                 
    def update_firestore(self):
    # Update data in Firestore
                    collection_ref = db.collection('student_info')
                    for i in range(self.table.rowCount()):
                        doc_data = {}
                        for j in range(self.table.columnCount()):
                            header = self.table.horizontalHeaderItem(j).text()
                            item = self.table.item(i, j)
                            if item is not None:
                                doc_data[header] = item.text()
                        if doc_data:
                            doc_ref = collection_ref.document()
                            doc_ref.set(doc_data)

    # Show success message
                    QMessageBox.information(self, "Success", "Data updated in Firestore.")            

    def select_columns(self):
    # Get selected columns
        selected_columns = []
        for checkbox in self.checkboxes:
            if checkbox.isChecked():
                selected_columns.append(checkbox.text())

    # Close the dialog
        self.close()

    # Do something with the selected columns
        if selected_columns:
            print("Selected columns:", selected_columns)
        # Read Excel file with selected columns
            df = pd.read_excel(self.file_name, usecols=selected_columns)

        # Create a QTableWidget and populate it with data
            self.table = QTableWidget()
            self.table.setColumnCount(len(selected_columns))
            self.table.setHorizontalHeaderLabels(selected_columns)
            self.table.setRowCount(len(df))
            for i, row in df.iterrows():
                for j, val in enumerate(row):
                    item = QTableWidgetItem(str(val))
                    item.setFlags(QtCore.Qt.ItemIsEditable | QtCore.Qt.ItemIsSelectable | QtCore.Qt.ItemIsEnabled)
                    self.table.setItem(i, j, item)

        # Create a dialog and show the table
            dialog = QDialog(self)
            layout = QVBoxLayout()
            layout.addWidget(self.table)

        # Add a submit button
            submit_button = QPushButton("Submit")
            submit_button.clicked.connect(self.update_firestore)
            layout.addWidget(submit_button)

            dialog.setLayout(layout)
            dialog.setWindowTitle("Selected Columns Data")
            dialog.setGeometry(200, 200, 500, 300)
            dialog.exec_()
        else:
            QMessageBox.warning(self, "Error", "Please select at least one column.")


if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    ex = ExcelImporter()

    # Set the checkboxes widget as the central widget for the main window
    central_widget = QWidget()
    ex.setCentralWidget(central_widget)
    central_layout = QVBoxLayout()
    central_layout.addWidget(ex.label)
    central_layout.addWidget(ex.button)
    central_widget.setLayout(central_layout)

    sys.exit(app.exec_())
import os
import pandas as pd
from tabulate import tabulate
from tkinter import Tk, messagebox
from PyQt5 import QtCore
from PyQt5.QtWidgets import QHeaderView, QApplication, QMainWindow, QWidget, QFormLayout, QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox, QCheckBox, QTableWidget, QTableWidgetItem, QVBoxLayout


class Table:
    def __init__(self, data):
        self.data = data

    def show(self):
        print(tabulate(self.data, headers="keys", tablefmt="fancy_grid"))


class KiaInvoicesTable(Table):
    def __init__(self, kia_data):
        kia_data = kia_data[['INVOICE #', 'DATE', 'SHIPMENT', 'CORE', 'SUBTOTAL', 'FREIGHT', 'HANDLING', 'MISC', 'RESTOCK FEE', 'TAX', 'INVOICE TOTAL']]
        kia_data.fillna("", inplace=True)
        kia_data = kia_data.astype({"INVOICE #": str})
        self.data = kia_data


def load_invoice():
    invoice_file = "extracted_data.xlsx"
    if not os.path.exists(invoice_file):
        QMessageBox.critical(None, "Error", f"{invoice_file} does not exist.")
        return
    invoice_data = pd.read_excel(invoice_file, sheet_name='Kia Invoices')
    if invoice_data.iloc[:, 11:16].isnull().all().all():
        kia_table = KiaInvoicesTable(invoice_data)
        print("Kia Invoices loaded successfully.")
        return kia_table
    else:
        QMessageBox.critical(None, "Error", f"{invoice_file} does not exist.")
        return None
    
class MoneyTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.text().replace('$', '').replace(',', '')) < float(other.text().replace('$', '').replace(',', ''))
        except ValueError:
            return super().__lt__(other)

class InvoiceSelection(QWidget):
    def __init__(self, table_data):
        super().__init__()

        # Set up the table
        table_layout = QVBoxLayout()
        table_label = QLabel("Kia Invoices")
        table_layout.addWidget(table_label)
        self.table_widget = QTableWidget()
        self.table_widget.setRowCount(len(table_data))
        self.table_widget.setColumnCount(len(table_data.columns) + 1)
        self.table_widget.setHorizontalHeaderLabels(table_data.columns.tolist() + ['Select'])

        for i in range(len(table_data)):
            for j in range(len(table_data.columns)):
                if table_data.columns[j] in ["CORE", "SUBTOTAL", "FREIGHT", "HANDLING", "MISC", "RESTOCK FEE", "TAX", "INVOICE TOTAL"]:
                    item = MoneyTableWidgetItem(f"${float(table_data.iloc[i, j] or 0):,.2f}")
                else:
                    item = QTableWidgetItem(str(table_data.iloc[i, j]))
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                self.table_widget.setItem(i, j, item)
            
        for i in range(len(table_data)):
            checkbox = QCheckBox()
            checkbox.setChecked(False)
            checkbox_cell = self.table_widget.model().index(i, len(table_data.columns))
            self.table_widget.setIndexWidget(checkbox_cell, checkbox)
            checkbox.setStyleSheet("QCheckBox::indicator { alignment: center; }") # Center the checkbox
                
        # Auto adjust column widths
        self.table_widget.resizeColumnsToContents()
        
        # Enable sorting
        self.table_widget.setSortingEnabled(True)
        
        # Add the checkbox and button
        self.checkbox = QCheckBox("Select all")
        self.checkbox.stateChanged.connect(self.select_all_rows)
        self.select_button = QPushButton("Select Invoice")
        buttons_layout = QVBoxLayout()
        buttons_layout.addWidget(self.checkbox)
        buttons_layout.addWidget(self.select_button)

        # Combine the layouts and set them
        layout = QVBoxLayout()
        layout.addLayout(table_layout)
        layout.addWidget(self.table_widget)
        layout.addLayout(buttons_layout)
        self.setLayout(layout)

        # Set up the window properties
        self.setWindowTitle("Kia Invoices")
        self.setGeometry(100, 100, 850, 800)
        
    def sort_table(self, index):
        self.table_widget.sortByColumn(index, QtCore.Qt.AscendingOrder)
        
    def select_all_rows(self, state):
        for i in range(self.table_widget.rowCount()):
            checkbox = self.table_widget.cellWidget(i, self.table_widget.columnCount() - 1)
            checkbox.setChecked(state)

def show_invoice():
    kia_table = load_invoice()
    if kia_table:
        app = QApplication([])
        invoice_widget = InvoiceSelection(kia_table.data)
        invoice_widget.show()
        app.exec_()
    else:
        print("No invoice data to display.")


if __name__ == "__main__":
    show_invoice()

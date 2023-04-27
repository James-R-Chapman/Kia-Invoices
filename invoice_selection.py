import os
import pandas as pd
from tabulate import tabulate
from tkinter import Tk, messagebox
from PyQt5 import QtCore
from PyQt5.QtWidgets import QHeaderView, QApplication, QMainWindow, QWidget, QFormLayout, QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox, QCheckBox, QTableWidget, QTableWidgetItem, QVBoxLayout
from Load_inv_data import load_invoice, Table, KiaInvoicesTable, PartsInvoicesTable, CreditsInvoicesTable
  
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
        table_label = QLabel("Invoices Without Accounts Added")
        table_layout.addWidget(table_label)
        self.table_widget = QTableWidget()
        self.table_widget.setRowCount(len(table_data))
        self.table_widget.setColumnCount(len(table_data.columns) + 2) # Add 2 columns for the dropdown and apply buttons
        self.table_widget.setHorizontalHeaderLabels(table_data.columns.tolist() + ['Account', 'Apply'])

        for i in range(len(table_data)):
            for j in range(len(table_data.columns)):
                if table_data.columns[j] in ["CORE", "SUBTOTAL", "FREIGHT", "HANDLING", "MISC", "RESTOCK FEE", "TAX", "INVOICE TOTAL"]:
                    item = MoneyTableWidgetItem(f"${float(table_data.iloc[i, j] or 0):,.2f}")
                else:
                    item = QTableWidgetItem(str(table_data.iloc[i, j]))
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                self.table_widget.setItem(i, j, item)

            # Add combobox item for account selection
            item = ComboBoxTableWidgetItem(['2410', '2430', '2431', '7186', '7194', '7196'], '2410')
            self.table_widget.setItem(i, len(table_data.columns), item)

            # Add apply button
            button = QPushButton('Apply')
            button.clicked.connect(lambda _, i=i: self.apply_account(i))
            button.setToolTip('Apply selected account to all parts in invoice')
            self.table_widget.setCellWidget(i, len(table_data.columns) + 1, button)
            
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
        self.setWindowTitle("Invoice Selection")
        self.setGeometry(100, 100, 850, 800)
        
        self.select_button.clicked.connect(self.select_invoice)

    def select_invoice(self):
        selected_rows = []
        for i in range(self.table_widget.rowCount()):
            checkbox = self.table_widget.cellWidget(i, self.table_widget.columnCount() - 1)
            if checkbox.isChecked():
                invoice_num = self.table_widget.item(i, 0).text()
                selected_rows.append(invoice_num)
        if selected_rows:
            self.selected_invoices = selected_rows
            self.accept()
        else:
            QMessageBox.critical(None, "Error", "Please select at least one invoice.")
        
    def sort_table(self, index):
        self.table_widget.sortByColumn(index, QtCore.Qt.AscendingOrder)
        
    def select_all_rows(self, state):
        for i in range(self.table_widget.rowCount()):
            checkbox = self.table_widget.cellWidget(i, self.table_widget.columnCount() - 1)
            checkbox.setChecked(state)


if invoice_widget.exec_():
    selected_invoices = invoice_widget.selected_invoices
    # Search for selected invoices in Parts Invoices data frame
    parts_invoices_selected = parts_invoice_data[parts_invoice_data['Invoice Number'].isin(selected_invoices)]
    parts_table = PartsInvoicesTable(parts_invoices_selected)
    # Search for selected invoices in Credits Invoices data frame
    credits_invoices_selected = credits_invoice_data[credits_invoice_data['Invoice Number'].isin(selected_invoices)]
    credits_table = CreditsInvoicesTable(credits_invoices_selected)
    # rest of the code


if __name__ == "__main__":
    app = QApplication([])
    table_data = pd.DataFrame(...)  # insert your table data here
    invoice_widget = InvoiceSelection(table_data)
    invoice_widget.show()
    app.exec_()


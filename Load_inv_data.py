import os
import pandas as pd
from tabulate import tabulate
from tkinter import Tk, messagebox
from PyQt5 import QtCore
from PyQt5.QtWidgets import QHeaderView, QApplication, QMainWindow, QWidget, QFormLayout, QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox, QCheckBox, QTableWidget, QTableWidgetItem, QVBoxLayout


class ComboBoxTableWidgetItem(QTableWidgetItem):
    def __init__(self, options, default_value):
        super().__init__()
        self.options = options
        self.combo_box = QComboBox()
        self.combo_box.addItems(options)
        self.combo_box.setCurrentText(default_value)
        self.setText(default_value)
        self.setTextAlignment(0x0004) # Align text to center
        self.setFlags(self.flags() & ~QtCore.Qt.ItemIsEditable) # Make cell read-only
        self.setSizeHint(self.combo_box.sizeHint()) # Set size hint to size of combo box
        self.setData(QtCore.Qt.UserRole, default_value)

    def setData(self, role, value):
        if role == QtCore.Qt.UserRole:
            self.combo_box.setCurrentText(value)
            self.setText(value)

    def data(self, role):
        if role == QtCore.Qt.UserRole:
            return self.combo_box.currentText()
        return super().data(role)

    def createEditor(self, parent):
        editor = self.combo_box
        editor.currentIndexChanged.connect(self.commitData)
        return editor


class Table:
    def __init__(self, data):
        self.data = data
    
    def show(self):
        print(tabulate(self.data, headers='keys', tablefmt='fancy_grid'))

class KiaInvoicesTable(Table):
    def __init__(self, kia_data):
        kia_data = kia_data[['INVOICE #', 'DATE', 'SHIPMENT', 'CORE', 'SUBTOTAL', 'FREIGHT', 'HANDLING', 'MISC', 'RESTOCK FEE', 'TAX', 'INVOICE TOTAL']]
        kia_data.fillna("", inplace=True)
        kia_data = kia_data.astype({"INVOICE #": str})
        self.data = kia_data


class PartsInvoicesTable(Table):
    def __init__(self, parts_invoice_data):
        parts_invoice_data = parts_invoice_data[['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Status', 'QTY Order', 'QTY Ship', 'B/O Cancel', 'Cost', 'Ext Price']]
        parts_invoice_data.fillna("", inplace=True)
        parts_invoice_data = parts_invoice_data.astype({"Invoice Number": str})
        self.data = parts_invoice_data.drop_duplicates(subset=['Invoice Number'])
    def add_column_selector(self):
        self.column_selector = QComboBox()
        # add column names to the combo box
        self.column_selector.addItems(['L', 'M', 'N', 'O', 'P', 'Q'])
        # set the current column as the default selection
        self.column_selector.setCurrentIndex(self.column)
        # add the combo box to the bottom of the table
        self.addWidget(self.column_selector, self.rowCount(), 0, 1, self.columnCount())

class CreditsInvoicesTable(Table):
    def __init__(self, credits_invoice_data):
        credits_invoice_data = credits_invoice_data[['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Type', 'QTY', 'ACT Code', 'Handling Credit', 'Cost', 'Ext Price']]
        credits_invoice_data.fillna("", inplace=True)
        credits_invoice_data = credits_invoice_data.astype({"Invoice Number": str})
        self.data = credits_invoice_data.drop_duplicates(subset=['Invoice Number'])
    def add_column_selector(self):
        self.column_selector = QComboBox()
        # add column names to the combo box
        self.column_selector.addItems(['L', 'M', 'N', 'O', 'P', 'Q'])
        # set the current column as the default selection
        self.column_selector.setCurrentIndex(self.column)
        # add the combo box to the bottom of the table
        self.addWidget(self.column_selector, self.rowCount(), 0, 1, self.columnCount())
        
def load_invoice():
    invoice_file = "extracted_data.xlsx"
    if not os.path.exists(invoice_file):
        QMessageBox.critical(None, "Error", f"{invoice_file} does not exist.")
        return None, None, None
    invoice_data = pd.read_excel(invoice_file, sheet_name='Kia Invoices')
    parts_invoice_data = pd.read_excel(invoice_file, sheet_name='Parts Invoices')
    credits_invoice_data = pd.read_excel(invoice_file, sheet_name='Credits Invoices')
    if invoice_data.iloc[:, 11:16].isnull().all().all():
        kia_table = KiaInvoicesTable(invoice_data)
        print("Kia Invoices loaded successfully.")
        return kia_table, parts_invoice_data, credits_invoice_data
    else:
        QMessageBox.critical(None, "Error", f"{invoice_file} does not exist.")
        return None, None, None

def show_invoice():
    kia_table, parts_invoice_data, credits_invoice_data = load_invoice()
    if kia_table:
        app = QApplication([])
        invoice_widget = InvoiceSelection(kia_table.data)
        if invoice_widget.exec_():
            selected_invoices = invoice_widget.selected_invoices
            # Search for selected invoices in Parts Invoices data frame
            parts_invoices_selected = parts_invoice_data[parts_invoice_data['Invoice #'].isin(selected_invoices)]
            # Search for selected invoices in Credits Invoices data frame
            credits_invoices_selected = credits_invoice_data[credits_invoice_data['Invoice #'].isin(selected_invoices)]
            # Merge the parts and credits invoices data frames
            merged_data = pd.concat([parts_invoices_selected, credits_invoices_selected])
            # Create the table widget with the merged data
            table = PartsTable(merged_data)
            table.add_column_selector() # add the column selector
            table.show()
            app.exec_()
        else:
            print("Invoice selection canceled.")
    else:
        print("No invoice data to display.")
        
def select_invoice(self):
    selected_cells = self.table_widget.selectedItems()
    selected_invoices = {cell.rowValues[0] for cell in selected_cells[::self.table_widget.columnCount()]}
    if selected_invoices:
        self.selected_invoices = selected_invoices
        self.accept()
    else:
        QMessageBox.critical(None, "Error", "Please select at least one invoice.")

def apply_changes():
    # Get the selected column from the column dropdown list
    selected_column = column_dropdown.currentText()

    # Create a dictionary that maps each dropdown option to the corresponding column name in both data frames
    column_mapping = {
        '2410': 'L',
        '2430': 'M',
        '2431': 'N',
        '7186': 'O',
        '7194': 'P',
        '7196': 'Q'
    }

    # Get the selected parts data from the PartsTable widget
    selected_parts = table_widget.getSelectedParts()
    selected_parts_df = pd.DataFrame(selected_parts, columns=['Invoice Number', 'Part Number', 'Quantity', 'Net Price', 'Ext Price'])

    # Update the 'Ext Price' column in the 'Parts Invoices' data frame with the 'Ext Price' values of the selected parts
    for invoice_num in selected_invoices:
        parts_invoice_data.loc[parts_invoice_data['Invoice Number'] == invoice_num, column_mapping[selected_column]] += selected_parts_df['Ext Price']

    # Update the 'Ext Price' column in the 'Credits Invoices' data frame with the 'Ext Price' values of the selected parts
    for invoice_num in selected_invoices:
        credits_invoice_data.loc[credits_invoice_data['Invoice Number'] == invoice_num, column_mapping[selected_column]] += selected_parts_df['Ext Price']

    # save the updated data frames to 'extracted_data.xlsx'
    with pd.ExcelWriter('extracted_data.xlsx') as writer:
        parts_invoice_data.to_excel(writer, sheet_name='Parts Invoices', index=False)
        credits_invoice_data.to_excel(writer, sheet_name='Credits Invoices', index=False)
    print("Data saved to extracted_data.xlsx")

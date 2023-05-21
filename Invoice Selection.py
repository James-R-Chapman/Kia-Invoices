import os
import pandas as pd
from tabulate import tabulate
from PyQt5 import QtCore
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFormLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QMessageBox, QCheckBox, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QDialog
)


class Table:
    def __init__(self, data):
        self.data = data

    def show(self):
        print(tabulate(self.data, headers="keys", tablefmt="fancy_grid"))


class KiaInvoicesTable(Table):
    def __init__(self, kia_data):
        kia_data = kia_data[['INVOICE #', 'DATE', 'SHIPMENT', 'CORE', 'SUBTOTAL', 'FREIGHT', 'HANDLING', 'MISC', 'RESTOCK FEE', 'TAX', 'INVOICE TOTAL']]
        kia_data = kia_data.astype({"INVOICE #": str})
        self.data = kia_data.copy().fillna("")


class InvoiceDetailsWindow(QDialog):
    def __init__(self, invoice_data, window_title):
        super().__init__()

        layout = QVBoxLayout(self)
        table_label = QLabel(window_title)
        layout.addWidget(table_label)

        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(len(invoice_data))
        self.table_widget.setColumnCount(len(invoice_data.columns) + 1)
        self.table_widget.setHorizontalHeaderLabels(invoice_data.columns.tolist() + ['Dropdown'])

        self.dropdowns = []

        for i in range(len(invoice_data)):
            for j in range(len(invoice_data.columns)):
                item = QTableWidgetItem(str(invoice_data.iloc[i, j]))
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                self.table_widget.setItem(i, j, item)

            dropdown = QComboBox()
            dropdown.addItem("")
            dropdown.addItems(['2410', '2430', '2431', '7186', '7194', '7196'])
            self.table_widget.setCellWidget(i, len(invoice_data.columns), dropdown)
            self.dropdowns.append(dropdown)

        self.table_widget.resizeColumnsToContents()

        layout.addWidget(self.table_widget)

        # Add the "Apply to All" button and account dropdown
        apply_all_layout = QHBoxLayout()
        self.apply_all_dropdown = QComboBox()
        self.apply_all_dropdown.addItem("")
        self.apply_all_dropdown.addItems(['2410', '2430', '2431', '7186', '7194', '7196'])
        self.apply_all_button = QPushButton("Apply to All")
        self.apply_all_button.clicked.connect(self.apply_to_all)
        apply_all_layout.addWidget(self.apply_all_dropdown)
        apply_all_layout.addWidget(self.apply_all_button)

        layout.addLayout(apply_all_layout)

        # Add the "Close Invoice" button
        close_button_layout = QHBoxLayout()
        self.close_button = QPushButton("Close Invoice")
        self.close_button.clicked.connect(self.close_invoice)
        close_button_layout.addWidget(self.close_button)

        layout.addLayout(close_button_layout)

        self.setWindowTitle(window_title)
        self.resize(self.table_widget.horizontalHeader().length() + 20, 600)

    def apply_to_all(self):
        selected_account = self.apply_all_dropdown.currentText()
        if selected_account:
            for dropdown in self.dropdowns:
                dropdown.setCurrentText(selected_account)

    def close_invoice(self):
        all_coded = all(dropdown.currentText() != "" for dropdown in self.dropdowns)
        if not all_coded:
            QMessageBox.warning(self, "Error", "Not all parts coded")
            return

        invoice_data_file = "invoice_data.xlsx"

        if os.path.exists(invoice_data_file):
            invoice_data = pd.read_excel(invoice_data_file, sheet_name=None)
        else:
            invoice_data = {}

        selected_account = self.apply_all_dropdown.currentText()
        for i in range(len(self.dropdowns)):
            invoice_row = []

            # Copy existing columns to invoice_row
            for j in range(11):
                invoice_row.append(self.table_widget.item(i, j).text())

            ext_price = self.table_widget.item(i, 10).text()
            if selected_account == '2410':
                invoice_row.extend([ext_price] * 5)
                invoice_row.append('')
            elif selected_account == '2430':
                invoice_row.extend([''] * 1)
                invoice_row.append(ext_price)
                invoice_row.extend([''] * 4)
            elif selected_account == '2431':
                invoice_row.extend([''] * 2)
                invoice_row.append(ext_price)
                invoice_row.extend([''] * 3)
            elif selected_account == '7186':
                invoice_row.extend([''] * 3)
                invoice_row.append(ext_price)
                invoice_row.extend([''] * 2)
            elif selected_account == '7194':
                invoice_row.extend([''] * 4)
                invoice_row.append(ext_price)
                invoice_row.extend([''] * 1)
            elif selected_account == '7196':
                invoice_row.extend([''] * 5)
                invoice_row.append(ext_price)
            else:
                invoice_row.extend([''] * 6)

            invoice_data.loc[i] = invoice_row

            if invoice_number.startswith('P'):
                sheet_name = 'Parts Invoices'
            elif invoice_number.startswith('C'):
                sheet_name = 'Credits Invoices'
            else:
                sheet_name = 'Kia Invoices'

            if sheet_name not in invoice_data:
                invoice_data[sheet_name] = pd.DataFrame(columns=['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number',
                                                                'Status', 'QTY Order', 'QTY Ship', 'B/O Cancel', 'Cost', 'Ext Price',
                                                                '2410', '2430', '2431', '7186', '7194', '7196'])

            invoice_data[sheet_name].loc[i] = invoice_row

        with pd.ExcelWriter(invoice_data_file, engine='xlsxwriter') as writer:
            for sheet_name, df in invoice_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Copy headers from extracted_data.xlsx
                extracted_data_file = "extracted_data.xlsx"
                if os.path.exists(extracted_data_file):
                    extracted_data = pd.read_excel(extracted_data_file, sheet_name=sheet_name)
                    header_row = extracted_data.columns.tolist()
                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]
                    for col_num, header in enumerate(header_row):
                        worksheet.write(0, col_num, header)

                # Add column headers for L-Q
                if sheet_name == 'Kia Invoices':
                    header_row = ['2410', '2430', '2431', '7186', '7194', '7196']
                    for col_num, header in enumerate(header_row, start=len(extracted_data.columns)):
                        worksheet.write(0, col_num, header)

        # Perform additional logic for closing the invoice
        # For example, open the next invoice or return to the Kia invoices window

        # Close the current invoice window
        self.accept()



class InvoiceSelection(QWidget):
    def __init__(self, table_data, open_invoice_details_dialog):
        super().__init__()

        self.open_invoice_details_dialog = open_invoice_details_dialog
        self.selected_invoices = set()  # Store selected invoice numbers

        # Set up the table
        table_layout = QVBoxLayout()
        table_label = QLabel("Kia Invoices")
        table_layout.addWidget(table_label)
        self.table_widget = QTableWidget()
        self.table_widget.setRowCount(len(table_data))
        self.table_widget.setColumnCount(len(table_data.columns) + 1)
        self.table_widget.setHorizontalHeaderLabels(table_data.columns.tolist() + ['  Select  '])

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
            checkbox.setStyleSheet("QCheckBox::indicator { alignment: center; }")  # Center the checkbox

        # Auto adjust column widths
        self.table_widget.resizeColumnsToContents()
        
        # Add extra width for the 'Select' header
        select_header = self.table_widget.horizontalHeaderItem(len(table_data.columns))
        select_header.setTextAlignment(QtCore.Qt.AlignCenter)
        select_header.setSizeHint(QtCore.QSize(100, 30))

        # Add the checkbox and button
        self.checkbox = QCheckBox("Select all")
        self.checkbox.stateChanged.connect(self.select_all_rows)
        self.select_button = QPushButton("Select Invoice")
        self.select_button.clicked.connect(self.open_invoice_details_dialog)
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
        self.setGeometry(100, 100, 950, 800)


    def sort_table(self, index):
        self.table_widget.sortByColumn(index, QtCore.Qt.AscendingOrder)

    def select_all_rows(self, state):
        for i in range(self.table_widget.rowCount()):
            checkbox = self.table_widget.cellWidget(i, self.table_widget.columnCount() - 1)
            checkbox.setChecked(state)


class MoneyTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            return float(self.text().replace('$', '').replace(',', '')) < float(other.text().replace('$', '').replace(',', ''))
        except ValueError:
            return super().__lt__(other)


class KiaInvoicesApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Kia Invoices")
        self.setGeometry(100, 100, 950, 800)

        kia_table = self.load_invoice()
        if kia_table:
            invoice_widget = InvoiceSelection(kia_table.data, self.open_invoice_details_dialog)
            self.setCentralWidget(invoice_widget)
        else:
            QMessageBox.critical(self, "Error", "No invoice data to display.")


    def load_invoice(self):
        invoice_file = "extracted_data.xlsx"
        if not os.path.exists(invoice_file):
            QMessageBox.critical(self, "Error", f"{invoice_file} does not exist.")
            return
        invoice_data = pd.read_excel(invoice_file, sheet_name='Kia Invoices')
        if invoice_data.iloc[:, 11:16].isnull().all().all():
            kia_table = KiaInvoicesTable(invoice_data)
            print("Kia Invoices loaded successfully.")
            return kia_table
        else:
            QMessageBox.critical(self, "Error", f"{invoice_file} does not exist.")
            return None

    def open_invoice_details_dialog(self):
        selected_invoices = set()  # Use a set to store unique invoice numbers
        for i in range(self.centralWidget().table_widget.rowCount()):
            checkbox = self.centralWidget().table_widget.cellWidget(i, self.centralWidget().table_widget.columnCount() - 1)
            if checkbox.isChecked():
                invoice_number = self.centralWidget().table_widget.item(i, 0).text()
                selected_invoices.add(invoice_number)

        if not selected_invoices:
            QMessageBox.warning(self, "No Invoice Selected", "Please select at least one invoice.")
            return

        app = QApplication.instance()
        if app is None:
            app = QApplication([])

        iterator = iter(selected_invoices)
        next_invoice = next(iterator)  # Get the first invoice from the set

        if next_invoice.startswith('P'):
            invoice_details = self.get_parts_invoice_details([next_invoice])  # Wrap invoice number in a list
            window = InvoiceDetailsWindow(invoice_details, "Parts Invoice")
            window.exec_()
        elif next_invoice.startswith('C'):
            invoice_details = self.get_credits_invoice_details([next_invoice])  # Wrap invoice number in a list
            window = InvoiceDetailsWindow(invoice_details, "Credits Invoice")
            window.exec_()
        else:
            QMessageBox.warning(self, "Invalid Selection", "Please select invoices starting with either 'P' or 'C'.")



    def get_parts_invoice_details(self, invoice_numbers):
        invoice_file = "extracted_data.xlsx"
        invoice_data = pd.read_excel(invoice_file, sheet_name='Parts Invoices')
        invoice_data = invoice_data[invoice_data['Invoice Number'].isin(invoice_numbers)]
        return invoice_data[['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Status', 'QTY Order', 'QTY Ship', 'B/O Cancel', 'Cost', 'Ext Price']]

    def get_credits_invoice_details(self, invoice_numbers):
        invoice_file = "extracted_data.xlsx"
        invoice_data = pd.read_excel(invoice_file, sheet_name='Credits Invoices')
        invoice_data = invoice_data[invoice_data['Invoice Number'].isin(invoice_numbers)]
        return invoice_data[['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Type', 'QTY', 'ACT Code', 'Handling Credit', 'Cost', 'Ext Price']]


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = KiaInvoicesApp()
    window.show()
    sys.exit(app.exec_())

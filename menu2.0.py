import os
import shutil  # used for file operations

import openpyxl
import pandas as pd
from PyQt5 import QtCore
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from tabulate import tabulate


class Table:
    def __init__(self, data):
        self.data = data

    def show(self):
        print(tabulate(self.data, headers="keys", tablefmt="fancy_grid"))


class KiaInvoicesTable(Table):
    def __init__(self, kia_data):
        kia_data = kia_data[
            [
                "INVOICE #",
                "DATE",
                "SHIPMENT",
                "CORE",
                "SUBTOTAL",
                "FREIGHT",
                "HANDLING",
                "MISC",
                "RESTOCK FEE",
                "TAX",
                "INVOICE TOTAL",
            ]
        ]
        kia_data.fillna("", inplace=True)
        kia_data = kia_data.astype({"INVOICE #": str})
        self.data = kia_data


def load_invoice():
    invoice_file = "extracted_data.xlsx"
    if not os.path.exists(invoice_file):
        QMessageBox.critical(None, "Error", f"{invoice_file} does not exist.")
        return
    shutil.copy("extracted_data.xlsx", "extracted_data_backup.xlsx")
    invoice_data = pd.read_excel(invoice_file, sheet_name="Kia Invoices")
    if invoice_data.iloc[:, 11:16].isnull().all().all():
        kia_table = KiaInvoicesTable(invoice_data)
        return kia_table
    else:
        QMessageBox.critical(None, "Error", f"{invoice_file} does not exist.")
        return None


class MoneyTableWidgetItem(QTableWidgetItem):
    def __init__(self, text):
        super().__init__(text.replace("nan", ""))
    
    def __lt__(self, other):
        try:
            return float(self.text().replace("$", "").replace(",", "")) < float(
                other.text().replace("$", "").replace(",", "")
            )
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
        self.table_widget.setHorizontalHeaderLabels(
            table_data.columns.tolist() + ["Select"]
        )

        for i in range(len(table_data)):
            for j in range(len(table_data.columns)):
                if table_data.columns[j] in [
                    "CORE",
                    "SUBTOTAL",
                    "FREIGHT",
                    "HANDLING",
                    "MISC",
                    "RESTOCK FEE",
                    "TAX",
                    "INVOICE TOTAL",
                ]:
                    item = MoneyTableWidgetItem(
                        f"${float(table_data.iloc[i, j] or 0):,.2f}"
                    )
                else:
                    item = QTableWidgetItem(str(table_data.iloc[i, j]))
                item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
                self.table_widget.setItem(i, j, item)

        for i in range(len(table_data)):
            checkbox = QCheckBox()
            checkbox.setChecked(False)
            checkbox_cell = self.table_widget.model().index(i, len(table_data.columns))
            self.table_widget.setIndexWidget(checkbox_cell, checkbox)
            checkbox.setStyleSheet(
                "QCheckBox::indicator { alignment: center; }"
            )  # Center the checkbox

        # Auto adjust column widths
        self.table_widget.resizeColumnsToContents()

        # Enable sorting
        self.table_widget.setSortingEnabled(True)

        # Add the checkbox and button
        self.checkbox = QCheckBox("Select all")
        self.checkbox.stateChanged.connect(self.select_all_rows)
        self.select_button = QPushButton("Select Invoice")
        self.select_button.clicked.connect(self.select_invoice_menu)
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
            checkbox = self.table_widget.cellWidget(
                i, self.table_widget.columnCount() - 1
            )
            checkbox.setChecked(state)

    def select_invoice_menu(self):
        selected_rows = []
        for i in range(self.table_widget.rowCount()):
            checkbox = self.table_widget.cellWidget(
                i, self.table_widget.columnCount() - 1
            )
            if checkbox.isChecked():
                selected_rows.append(i)

        if not selected_rows:
            QMessageBox.critical(
                None, "Error", "Please select one or more invoices to proceed."
            )
            return

        for selected_row in selected_rows:
            invoice_num = self.table_widget.item(selected_row, 0).text()
            if invoice_num.startswith("P"):
                self.select_invoice_p(invoice_num)
            elif invoice_num.startswith("C"):
                self.select_invoice_c(invoice_num)
            else:
                QMessageBox.critical(
                    None,
                    "Error",
                    f"The selected invoice ({invoice_num}) is not a 'P' or 'C' invoice.",
                )

    def select_invoice_p(self, invoice_num):
        invoice_data = pd.read_excel("extracted_data.xlsx", sheet_name="Parts Invoices")
        invoice_data = invoice_data[invoice_data["Invoice Number"] == invoice_num]
        invoice_type = "P"
        # Define the columns to display in the pop-up window
        columns = [
            "Invoice Number",
            "Invoice Date",
            "Line #",
            "Order #",
            "Part Number",
            "Status",
            "QTY Order",
            "QTY Ship",
            "B/O Cancel",
            "Cost",
            "Ext Price",
        ]

        # Create the table widget with the account dropdowns and apply buttons
        table = create_invoice_table_widget(invoice_data, columns, invoice_num, invoice_type)
        save_to_dropdown = table.findChild(QComboBox, "save_to_dropdown")
        print(save_to_dropdown)  # add this line

        # Connect the "Apply All" button to the apply_all_accounts function
        apply_all_button = table.findChild(QPushButton, "apply_all_button")
        apply_all_button.clicked.connect(
            lambda: apply_all_accounts(table, save_to_dropdown)
        )

    def select_invoice_c(self, invoice_num):
        invoice_data = pd.read_excel("extracted_data.xlsx", sheet_name="Credits Invoices")
        invoice_data = invoice_data[invoice_data["Invoice Number"] == invoice_num]
        invoice_type = "C"
        # Define the columns to display in the pop-up window
        columns = [
            "Invoice Number",
            "Invoice Date",
            "Line #",
            "Order #",
            "Part Number",
            "Type",
            "QTY",
            "ACT Code",
            "Handling Credit",
            "Cost",
            "Ext Price",
        ]

        # Create the table widget with the account dropdowns and apply buttons
        table = create_invoice_table_widget(invoice_data, columns, invoice_num, invoice_type)
        save_to_dropdown = table.findChild(QComboBox, "save_to_dropdown")
        print(save_to_dropdown)  # add this line  

        # Connect the "Apply All" button to the apply_all_accounts function
        apply_all_button = table.findChild(QPushButton, "apply_all_button")
        apply_all_button.clicked.connect(lambda: apply_all_accounts(table, save_to_dropdown))

def create_invoice_table_widget(invoice_data, columns, invoice_num, invoice_type):
    # Create the table widget
    table = QTableWidget()
    table.setRowCount(len(invoice_data))
    table.setColumnCount(len(columns) + 1)
    table.setHorizontalHeaderLabels(columns + ["Account"])

    # Add the dropdown menu for the account selection
    save_to_dropdown = QComboBox()
    save_to_dropdown.addItems(["", "2410", "2430", "2431", "7186", "7194", "7196"])

    # Create account dropdowns for each row
    for i in range(len(invoice_data)):
        for j in range(len(columns)):
            if columns[j] in ["Cost", "Ext Price"]:
                item = MoneyTableWidgetItem(
                    f"${float(invoice_data.iloc[i, j] or 0):,.2f}"
                )
            else:
                item = QTableWidgetItem(str(invoice_data.iloc[i, j]))
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsEditable)
            table.setItem(i, j, item)

        account_dropdown = QComboBox()
        account_dropdown.addItems(["", "2410", "2430", "2431", "7186", "7194", "7196"])
        account_dropdown_widget = QWidget()
        account_dropdown_layout = QHBoxLayout(account_dropdown_widget)
        account_dropdown_layout.addWidget(account_dropdown)
        account_dropdown_layout.setContentsMargins(0, 0, 0, 0)
        table.setCellWidget(i, len(columns), account_dropdown_widget)

    # Create the 'Apply All' button and the dropdown for the new .xlsx file
    apply_button_layout = QHBoxLayout()
    apply_all_button = QPushButton("Apply All")
    save_to_label = QLabel("Apply all parts to the following account:")
    save_button = QPushButton("Save")
    apply_button_layout.addWidget(save_to_label)
    apply_button_layout.addWidget(save_to_dropdown)
    apply_button_layout.addWidget(apply_all_button)
    apply_button_layout.addWidget(save_button)

    # Define the apply_all_handler function
    def apply_all_handler():
        apply_all_accounts(table, save_to_dropdown)

    # Connect the 'Apply All' button to the apply_all_handler function
    apply_all_button.clicked.connect(lambda: apply_all_handler(table, save_to_dropdown))
    save_button.clicked.connect(lambda _, table=table, invoice_data=invoice_data, invoice_type=invoice_type: save_handler(invoice_data, table, invoice_type))


    # Create the layout for the window
    popup_layout = QVBoxLayout()
    popup_layout.addWidget(QLabel(f"Invoice Number: {invoice_num}"))
    popup_layout.addWidget(table)
    popup_layout.addLayout(apply_button_layout)

    # Create the window
    popup = QDialog()
    popup.setWindowTitle("Invoice Details")
    popup.setLayout(popup_layout)

    # Resize the columns based on the contents
    table.resizeColumnsToContents()

    # Set the properties of the window and show it
    popup.setGeometry(100, 100, 800, 600)
    popup.exec_()

    return popup

def show_invoice():
    kia_table = load_invoice()
    if kia_table:
        app = QApplication([])
        invoice_widget = InvoiceSelection(kia_table.data)
        invoice_widget.show()
        app.exec_()
    else:
        print("No invoice data to display.")
        
def save_handler(invoice_data, table, invoice_type):
    # Get the selected account from the dropdown
    for i in range(table.rowCount()):
        account_dropdown_widget = table.cellWidget(i, table.columnCount() - 1)
        account = account_dropdown_widget.findChild(QComboBox).currentText()
        
        # Determine the target sheet based on the invoice type
        if invoice_type.startswith("P"):
            target_sheet = "Parts Invoices"
        elif invoice_type.startswith("C"):
            target_sheet = "Credits Invoices"
        else:
            raise ValueError(f"Invalid invoice type: {invoice_type}")
        
        # Write the account to the appropriate cell in the spreadsheet
        row = table.item(i, 0).text()
        col = "L"  # fixed value
        cell = f"{col}{i + 2}"
        sheet = None  # initialize sheet variable
        for j in range(1, workbook[target_sheet].max_row + 1):
            if workbook[target_sheet][f"A{j}"].value == row:
                sheet = workbook[target_sheet]
                sheet[cell] = account
                break

        # Raise an error if the row type is invalid
        if not sheet:
            raise ValueError(f"Invalid row type: {row}")
        

    # Save the .xlsx file
    workbook.save("data_extracted_backup.xlsx")
    print("Invoice saved successfully.")



if __name__ == "__main__":
    show_invoice()

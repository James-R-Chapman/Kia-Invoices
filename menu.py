import os
import pandas as pd
from tabulate import tabulate
from tkinter import Tk, messagebox
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QFormLayout, QLabel, QLineEdit, QComboBox, QPushButton, QMessageBox



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


def show_invoice():
    kia_table = load_invoice()
    if kia_table:
        app = QApplication([])
        root = Tk()
        root.withdraw()
        QMessageBox.information(None, "Invoice Data", tabulate(kia_table.data, headers="keys", tablefmt="fancy_grid"))
        app.exec_()
    else:
        print("No invoice data to display.")


if __name__ == "__main__":
    show_invoice()

import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QCheckBox, QComboBox, QTableWidget, QTableWidgetItem
from PyQt5.QtWidgets import QDialog, QDialogButtonBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItem, QStandardItemModel
import pandas as pd
import datetime

class InvoicePartDialog(QMainWindow):

    def __init__(self, invoice_file, credit_file):
        super().__init__()
        self.invoice_file = invoice_file
        self.credit_file = credit_file

        # Load invoice and credit data
        self.invoice_data = pd.read_excel(invoice_file, sheet_name='Parts Invoices', usecols=[0,1,2,3,4,5,6,7,8,9,10])
        self.invoice_data.columns = ['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Status', 'QTY Order', 'QTY Ship', 'B/O Cancel', 'Cost', 'Ext Price']
        self.invoice_data.insert(loc=len(self.invoice_data.columns), column='Account', value='')

        self.credit_data = pd.read_excel(credit_file, sheet_name='Credits Invoices', usecols=[0,1,2,3,4,5,6,7,8,9,10])
        self.credit_data.columns = ['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Type', 'QTY', 'ACT Code', 'Handling Credit', 'Cost', 'Ext Price']
        self.credit_data.insert(loc=len(self.credit_data.columns), column='Account', value='')

        # Create the table widget
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(['Invoice Number', 'Invoice Date', 'Order #', 'Part Number', 'Qty Ship', 'Cost', 'Ext Price'])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)

        # Add the data to the table
        self.add_invoice_data_to_table()
        self.add_credit_data_to_table()

        # Add a checkbox to select all rows
        self.select_all_checkbox = QCheckBox('Select All')
        self.select_all_checkbox.stateChanged.connect(self.select_all)

        # Add a combobox to select the account
        self.account_combobox = QComboBox()
        self.account_combobox.addItems(['2410', '2430', '2431', '7186', '7194', '7196'])

        # Add a checkbox to select all accounts
        self.select_all_accounts_checkbox = QCheckBox('Select All Accounts')
        self.select_all_accounts_checkbox.stateChanged.connect(self.select_all_accounts)

        # Add a button to apply the account to the selected rows
        self.apply_button = QPushButton('Apply')
        self.apply_button.clicked.connect(self.apply_account)

        # Create the layout
        vbox = QVBoxLayout()
        vbox.addWidget(self.table)
        hbox1 = QHBoxLayout()
        hbox1.addWidget(self.select_all_checkbox)
        hbox1.addWidget(QLabel('Account:'))
        hbox1.addWidget(self.account_combobox)
        hbox1.addWidget(self.apply_button)
        vbox.addLayout(hbox1)
        hbox2 = QHBoxLayout()
        hbox2.addWidget(self.select_all_accounts_checkbox)
        vbox.addLayout(hbox2)
        widget = QWidget()
        widget.setLayout(vbox)
        self.setCentralWidget(widget)

    def add_invoice_data_to_table(self):
        # Load the invoice data
        invoice_data = pd.read_excel(self.invoice_file, sheet_name='Parts Invoices', usecols=[0,1,2,3,4,5,6,7,8,9,10])
        invoice_data.columns = ['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Status', 'QTY Order', 'QTY Ship', 'B/O Cancel', 'Cost', 'Ext Price']
        invoice_data = invoice_data.sort_values(by=['Invoice Number', 'Part Number'])

        # Add the data to the table
        current_invoice = None
        for index, row in invoice_data.iterrows():
            invoice_number = row['Invoice Number']
            if current_invoice != invoice_number:
                # Add a new row for the invoice number
                self.table.insertRow(self.table.rowCount())
                current_invoice = invoice_number
            else:
                invoice_parts = invoice_data[invoice_data['Invoice Number'] == invoice_number]
                invoice_date = datetime.datetime.strptime(invoice_parts.iloc[0]['Invoice Date'], '%m/%d/%y').strftime('%Y-%m-%d')
                row_position = self.table.rowCount()
                if invoice_number.startswith('P'):
                    self.table.setHorizontalHeaderLabels(['Invoice Number', 'Invoice Date', 'Order #', 'Part Number', 'Qty Ship', 'Cost', 'Ext Price', 'Acct #'])
                    for index, row in invoice_parts.iterrows():
                        order_number = str(row['Order #'])
                        part_number = str(row['Part Number'])
                        qty_ship = str(row['QTY Ship'])
                        cost = str(row['Cost'])
                        ext_price = str(row['Ext Price'])
                        self.table.insertRow(row_position)
                        self.table.setItem(row_position, 0, QTableWidgetItem(str(invoice_number)))
                        self.table.setItem(row_position, 1, QTableWidgetItem(invoice_date))
                        self.table.setItem(row_position, 2, QTableWidgetItem(order_number))
                        self.table.setItem(row_position, 3, QTableWidgetItem(part_number))
                        self.table.setItem(row_position, 4, QTableWidgetItem(qty_ship))
                        self.table.setItem(row_position, 5, QTableWidgetItem(cost))
                        self.table.setItem(row_position, 6, QTableWidgetItem(str(ext_price)))
                        checkbox = QCheckBox()
                        checkbox.setObjectName(f"{invoice_number}|{part_number}")
                        self.table.setCellWidget(row_position, 7, checkbox)
                        combobox = QComboBox()
                        combobox.setObjectName(f"{invoice_number}|{part_number}")
                        combobox.addItems(['', '2410', '2430', '2431', '7186', '7194', '7196'])
                        self.table.setCellWidget(row_position, 8, combobox)
                    checkbox = QCheckBox()
                    checkbox.setObjectName(f"{invoice_number}|All Parts")
                    self.table.setCellWidget(row_position + 1, 7, checkbox)
                    checkbox.stateChanged.connect(lambda state, invoice_number=invoice_number: self.show_part_picker_dialog(invoice_number))
                    apply_button = QPushButton('Apply')
                    apply_button.setObjectName(f"{invoice_number}|All Parts")
                    apply_button.clicked.connect(lambda state, invoice_number=invoice_number: self.apply_account_to_invoice(invoice_number))
                    self.table.setCellWidget(row_position + 1, 8, apply_button)
                else:
                    self.table.setHorizontalHeaderLabels(['Invoice Number', 'Invoice Date', 'Type', 'QTY', 'ACT Code', 'Cost', 'Ext Price', 'Acct #']) # added 'Acct #' to the header labels
                    for index, row in invoice_parts.iterrows():
                        qty = str(row['QTY'])
                        act_code = str(row['ACT Code'])
                        cost = str(row['Cost'])
                        ext_price = str(row['Ext Price'])
                        self.table.insertRow(row_position)
                        self.table.setItem(row_position, 0, QTableWidgetItem(str(invoice_number)))
                        self.table.setItem(row_position, 1, QTableWidgetItem(invoice_date))
                        self.table.setItem(row_position, 2, QTableWidgetItem(row['Type']))
                        self.table.setItem(row_position, 3, QTableWidgetItem(qty))
                        self.table.setItem(row_position, 4, QTableWidgetItem(act_code))
                        self.table.setItem(row_position, 5, QTableWidgetItem(cost))
                        self.table.setItem(row_position, 6, QTableWidgetItem(str(ext_price)))
                        acct_number = QComboBox() # added QComboBox to display account numbers
                        acct_number.addItems(['', '2410', '2430', '2431', '7186', '7194', '7196']) # added account number options
                        acct_number.currentIndexChanged.connect(lambda state, row_position=row_position, col=7, invoice_number=invoice_number, part_number=act_code: self.set_acct_number(row_position, col, invoice_number, part_number, acct_number.currentText())) # added currentIndexChanged to change the account number value when the combo box is changed
                        self.table.setCellWidget(row_position, 8, acct_number) # added the combo box to the table


    
    def add_credit_data_to_table(self):
        # Load the credit data
        credit_data = pd.read_excel(self.credit_file, sheet_name='Credits Invoices', usecols=[0,1,2,3,4,5,6,7,8,9,10])
        credit_data.columns = ['Invoice Number', 'Invoice Date', 'Line #', 'Order #', 'Part Number', 'Type', 'QTY', 'ACT Code', 'Handling Credit', 'Cost', 'Ext Price']
        credit_data = credit_data.sort_values(by=['Invoice Number', 'Part Number'])

        # Add the data to the table
        current_invoice = None
        for index, row in credit_data.iterrows():
            invoice_number = row['Invoice Number']
            if current_invoice != invoice_number:
                # Add a new row for the invoice number
                self.table.insertRow(self.table.rowCount())
                current_invoice = invoice_number
            else:
                invoice_parts = credit_data[credit_data['Invoice Number'] == invoice_number]
                invoice_date = datetime.datetime.strptime(invoice_parts.iloc[0]['Invoice Date'], '%m/%d/%y').strftime('%Y-%m-%d')
                row_position = self.table.rowCount()
                if invoice_number.startswith('P'):
                    self.table.setHorizontalHeaderLabels(['Invoice Number', 'Invoice Date', 'Order #', 'Part Number', 'Qty Ship', 'Cost', 'Ext Price'])
                    for index, row in invoice_parts.iterrows():
                        order_number = str(row['Order #'])
                        part_number = str(row['Part Number'])
                        qty_ship = str(row['QTY Ship'])
                        cost = str(row['Cost'])
                        ext_price = str(row['Ext Price'])
                        self.table.insertRow(row_position)
                        self.table.setItem(row_position, 0, QTableWidgetItem(str(invoice_number)))
                        self.table.setItem(row_position, 1, QTableWidgetItem(invoice_date))
                        self.table.setItem(row_position, 2, QTableWidgetItem(order_number))
                        self.table.setItem(row_position, 3, QTableWidgetItem(part_number))
                        self.table.setItem(row_position, 4, QTableWidgetItem(qty_ship))
                        self.table.setItem(row_position, 5, QTableWidgetItem(cost))
                        self.table.setItem(row_position, 6, QTableWidgetItem(str(ext_price)))
                        checkbox = QCheckBox()
                        checkbox.setObjectName(f"{invoice_number}|{part_number}") 
                        self.table.setCellWidget(row_position, 7, checkbox)
                    checkbox = QCheckBox()
                    checkbox.setObjectName(f"{invoice_number}|All Parts") 
                    self.table.setCellWidget(row_position + 1, 7, checkbox)
                    checkbox.stateChanged.connect(lambda state, invoice_number=invoice_number: self.show_part_picker_dialog(invoice_number))
                else:
                    self.table.setHorizontalHeaderLabels(['Invoice Number', 'Invoice Date', 'Type', 'QTY', 'ACT Code', 'Cost', 'Ext Price'])
                    for index, row in invoice_parts.iterrows():
                        qty = str(row['QTY'])
                        act_code = str(row['ACT Code'])
                        cost = str(row['Cost'])
                        ext_price = str(row['Ext Price'])
                        self.table.insertRow(row_position)
                        self.table.setItem(row_position, 0, QTableWidgetItem(str(invoice_number)))
                        self.table.setItem(row_position, 1, QTableWidgetItem(invoice_date))
                        self.table.setItem(row_position, 2, QTableWidgetItem(row['Type']))
                        self.table.setItem(row_position, 3, QTableWidgetItem(qty))
                        self.table.setItem(row_position, 4, QTableWidgetItem(act_code))
                        self.table.setItem(row_position, 5, QTableWidgetItem(cost))
                        self.table.setItem(row_position, 6, QTableWidgetItem(str(ext_price)))
                        checkbox = QCheckBox()
                        checkbox.setObjectName(f"{invoice_number}|{act_code}")
                        self.table.setCellWidget(row_position, 7, checkbox)
                        checkbox = QCheckBox()
                        checkbox.setObjectName(f"{invoice_number}|All ACT Codes")
                        self.table.setCellWidget(row_position + 1, 7, checkbox)
                        checkbox.stateChanged.connect(lambda state, invoice_number=invoice_number: self.show_act_code_picker_dialog(invoice_number))

    def show_part_picker_dialog(self, invoice_number):
        # Create the dialog
        dialog = QDialog(self)
        dialog.setWindowTitle('Select Parts')
        dialog.setModal(True)

        # Create the table widget
        table = QTableWidget()
        table.setColumnCount(1)
        table.setHorizontalHeaderLabels(['Part Number'])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)

        # Get the invoice parts
        invoice_parts = self.invoice_data[self.invoice_data['Invoice Number'] == invoice_number]
        parts = invoice_parts['Part Number'].unique()

        # Add the parts to the table
        for part_number in parts:
            table.insertRow(table.rowCount())
            table.setItem(table.rowCount()-1, 0, QTableWidgetItem(part_number))

        # Create the layout
        vbox = QVBoxLayout()
        vbox.addWidget(table)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        vbox.addWidget(button_box)
        dialog.setLayout(vbox)

        # Show the dialog
        if dialog.exec_() == QDialog.Accepted:
            # Apply the account to the selected parts
            rows = set()
            for index in range(self.table.rowCount()):
                checkbox = self.table.cellWidget(index, 7)
                if checkbox and checkbox.isChecked():
                    invoice_item = self.table.item(index, 0)
                    part_item = self.table.item(index, 4)
                    if invoice_item and part_item:
                        if invoice_item.text() == invoice_number and part_item.text() in parts:
                            rows.add(index)
            acct_number = self.account_combobox.currentText()
            for row in rows:
                acct_number_item = self.table.item(row, 8) # changed column to 8 to match the newly added account number column
                if acct_number_item:
                    acct_number_item.setText(acct_number)
                    invoice_item = self.table.item(row, 0)
                    invoice_number = invoice_item.text()
                    part_number_item = self.table.item(row, 4)
                    part_number = part_number_item.text()
                    self.invoice_data.loc[(self.invoice_data['Invoice Number'] == invoice_number) & (self.invoice_data['Part Number'] == part_number), 'Account'] = acct_number # added line to update the invoice_data DataFrame with the account number

            self.invoice_data.to_excel('K_inv.xlsx', index=False) # changed filename to 'K_inv.xlsx'
            self.close()


    def apply_account_to_parts(self, table, account):
        totals_by_invoice = {}
        for row in range(table.rowCount()):
            select_checkbox = table.cellWidget(row, 0)
            if select_checkbox is not None and select_checkbox.isChecked():
                invoice_number_item = table.item(row, 0)
                invoice_number = invoice_number_item.text()
                account_item = table.item(row, 4)
                account_item.setText(account)
                price_item = table.item(row, 3)
                price = float(price_item.text())
                invoice_total = totals_by_invoice.get(invoice_number, 0)
                totals_by_invoice[invoice_number] = invoice_total + price
        self.update_kia_invoices(totals_by_invoice)


    def select_all(self, table, state):
        for row in range(table.rowCount()):
            select_checkbox = table.cellWidget(row, 0)
            if select_checkbox is not None:
                select_checkbox.setChecked(state)
                
    def apply_account(self):
        rows = set()
        for index in range(self.table.rowCount()):
            checkbox = self.table.cellWidget(index, 7)
            if checkbox and checkbox.isChecked():
                rows.add(index)
        acct_number = self.account_combobox.currentText()
        for row in rows:
            part_number_item = self.table.item(row, 4)
            acct_number_item = self.table.item(row, 8) # changed column to 8 to match the newly added account number column
            if part_number_item and acct_number_item:
                part_number = part_number_item.text()
                acct_number_item.setText(acct_number)
                invoice_item = self.table.item(row, 0)
                invoice_number = invoice_item.text()
                self.invoice_data.loc[(self.invoice_data['Invoice Number'] == invoice_number) & (self.invoice_data['Part Number'] == part_number), 'Account'] = acct_number # added line to update the invoice_data DataFrame with the account number
        self.invoice_data.to_excel('K_inv.xlsx', index=False) # changed filename to 'K_inv.xlsx'
        self.close()

    
    def select_all_accounts(self, state):
        if state == Qt.Checked:
            # Select all accounts
            for row in range(self.table.rowCount()):
                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                checkbox_item.setCheckState(Qt.Checked)
                self.table.setItem(row, 11, checkbox_item)
        else:
            # Deselect all accounts
            for row in range(self.table.rowCount()):
                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                checkbox_item.setCheckState(Qt.Unchecked)
                self.table.setItem(row, 11, checkbox_item)

    def apply_account_to_parts(self, table, account):
        totals_by_invoice = {}
        for row in range(table.rowCount()):
            select_checkbox = table.cellWidget(row, 0)
            if select_checkbox is not None and select_checkbox.isChecked():
                invoice_number_item = table.item(row, 0)
                invoice_number = invoice_number_item.text()
                account_item = table.item(row, 4)
                account_item.setText(account)
                price_item = table.item(row, 3)
                price = float(price_item.text())
                invoice_total = totals_by_invoice.get(invoice_number, 0)
                totals_by_invoice[invoice_number] = invoice_total + price
        self.update_kia_invoices(totals_by_invoice)
        
    def update_kia_invoices(self, totals_by_invoice):
        for row in range(self.kia_invoices_table.rowCount()):
            invoice_item = self.kia_invoices_table.item(row, 0)
            if invoice_item is not None:
                invoice_number = invoice_item.text()
                if invoice_number in totals_by_invoice:
                    for col in range(11, self.kia_invoices_table.columnCount()):
                        account_item = self.kia_invoices_table.item(row, col)
                        if account_item is not None:
                            account = account_item.text()
                            if account in self.account_numbers:
                                total = totals_by_invoice[invoice_number]
                                account_total = float(account_item.text()) if account_item.text() != '' else 0
                                self.kia_invoices_table.setItem(row, col, QTableWidgetItem(str(account_total + total)))
                    del totals_by_invoice[invoice_number]

    def update_kia_invoices(self, total_by_account):
        # Load Kia invoice data
        kia_invoice_data = pd.read_excel('Kia Invoices.xlsx', usecols=[0, 4, 5, 6, 7, 8, 9, 10], dtype={'INVOICE #': str})
        kia_invoice_data.columns = ['INVOICE #', '2410', '2430', '2431', '7186', '7194', '7196', 'INVOICE TOTAL']

        # Create a copy of the data with the updated account totals
        updated_data = kia_invoice_data.copy()
        for invoice_num, accounts in total_by_account.items():
            for account, data in accounts.items():
                account_col = str(account)
                updated_data.loc[updated_data['INVOICE #'] == invoice_num, account_col] += data['Total']

        # Save the updated data to a new file
        writer = pd.ExcelWriter('Kia Invoices (Updated).xlsx', engine='xlsxwriter')
        updated_data.to_excel(writer, index=False)
        writer.save()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    dialog = InvoicePartDialog('extracted_data.xlsx', 'extracted_data.xlsx')
    dialog.show()
    sys.exit(app.exec_())



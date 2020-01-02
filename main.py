import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtSerialPort import *
import secrets
import csv
import scanner, ag_scanner
import config
import re


class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Alpha Matrix")
        self.setWindowIcon(QIcon('icons/qr-code-32.png'))
        self.setGeometry(450, 150, 1350, 750)
        # self.setFixedSize(self.size())
        self.UI()
        self.show()
        # self.signal_connect()

    def UI(self):
        self.toolbar()
        self.tabWidget()
        self.widgets()
        self.layouts()

    def toolbar(self):
        self.toolbar = self.addToolBar('Toolbar')
        self.toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)

        self.generateCodeAction = QAction(QIcon('icons/process-32.png'), 'Gen Codes', self)
        self.toolbar.addAction(self.generateCodeAction)
        self.generateSerialsAction = QAction(QIcon('icons/padlock-64.png'), 'Gen Serials', self)
        self.toolbar.addAction(self.generateSerialsAction)
        self.toolbar.addSeparator()
        self.setScannerAction = QAction(QIcon('icons/barcode-scanner-32'), 'Set Scanner', self)
        self.setScannerAction.triggered.connect(self.setScanner)
        self.toolbar.addAction(self.setScannerAction)
        self.setAgScanner = QAction(QIcon('icons/qr-code-scan-24'), "Set Aggregate", self)
        self.setAgScanner.triggered.connect(self.AgScanner)
        self.toolbar.addAction(self.setAgScanner)
        self.setRejectorAction = QAction(QIcon('icons/trash.png'), "Set Rejector", self)
        self.toolbar.addAction(self.setRejectorAction)
        self.toolbar.addSeparator()

    def tabWidget(self):
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        self.tab4 = QWidget()
        self.tabs.addTab(self.tab1, 'Printing')
        self.tabs.addTab(self.tab2, 'Scanning')
        self.tabs.addTab(self.tab3, 'Aggregation')
        self.tabs.addTab(self.tab4, 'Statistics')

    def widgets(self):
        # Main Widgets
        self.codesTable = QTableWidget()
        self.codesTable.setColumnCount(6)
        self.codesTable.setHorizontalHeaderItem(0, QTableWidgetItem("NIE"))
        self.codesTable.setHorizontalHeaderItem(1, QTableWidgetItem("NIE Exp"))
        self.codesTable.setHorizontalHeaderItem(2, QTableWidgetItem("Batch"))
        self.codesTable.setHorizontalHeaderItem(3, QTableWidgetItem("Prod Date"))
        self.codesTable.setHorizontalHeaderItem(4, QTableWidgetItem("Expired"))
        self.codesTable.setHorizontalHeaderItem(5, QTableWidgetItem("Serial"))
        self.codesTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.codesTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.codesTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)

        # right top widgets
        self.nieEntry = QLineEdit()
        self.nieExpiryEntry = QDateEdit()
        self.nieExpiryEntry.setCalendarPopup(True)
        self.nieExpiryEntry.setDate(QDate.currentDate())
        self.batchEntry = QLineEdit()
        self.productionDateEntry = QDateEdit()
        self.productionDateEntry.setDate(QDate.currentDate())
        self.productionDateEntry.setCalendarPopup(True)
        self.expiryDateEntry = QDateEdit()
        self.expiryDateEntry.setCalendarPopup(True)
        self.expiryDateEntry.setDate(QDate.currentDate())
        self.codeQuantityEntry = QLineEdit()
        self.codeQuantityEntry.setMaxLength(8)
        self.codeQuantityEntry.setMaximumWidth(100)
        self.onlyInt = QIntValidator()
        self.codeQuantityEntry.setValidator(self.onlyInt)
        self.generateButton = QPushButton("Generate")
        self.generateButton.clicked.connect(self.generateCodes)
        self.resetButton = QPushButton("Reset")
        self.resetButton.clicked.connect(self.resetCodeGen)
        self.exportButton = QPushButton("Export CSV")
        self.exportButton.setDisabled(True)
        self.exportButton.clicked.connect(self.export_to_csv)

        # Right Mid Widgets
        self.serialQuantityEntry = QLineEdit()
        self.serialQuantityEntry.setMaxLength(8)
        self.serialQuantityEntry.setMaximumWidth(100)
        self.serialQuantityEntry.textChanged.connect(self.serial_quantity_changed)
        self.serialQuantityEntry.setValidator(self.onlyInt)
        self.serialGenExportButton = QPushButton("Generate + Export")
        self.serialGenExportButton.setEnabled(False)
        self.serialGenExportButton.clicked.connect(self.serial_gen_export)

        # Scanner Widgets
        self.scannedCodesTable = QTableWidget()
        self.scannedCodesTable.setColumnCount(6)
        self.scannedCodesTable.setHorizontalHeaderItem(0, QTableWidgetItem("NIE"))
        self.scannedCodesTable.setHorizontalHeaderItem(1, QTableWidgetItem("NIE Exp"))
        self.scannedCodesTable.setHorizontalHeaderItem(2, QTableWidgetItem("Batch"))
        self.scannedCodesTable.setHorizontalHeaderItem(3, QTableWidgetItem("Prod Date"))
        self.scannedCodesTable.setHorizontalHeaderItem(4, QTableWidgetItem("Expired"))
        self.scannedCodesTable.setHorizontalHeaderItem(5, QTableWidgetItem("Serial"))
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)

        # Scan top right widgets
        self.scan_nie_label = QLabel("NIE : ")
        self.scan_nie_text = QLineEdit()
        self.scan_nie_text.setReadOnly(True)
        self.scan_nie_exp_label = QLabel('NIE Expired : ')
        self.scan_nie_exp_text = QLineEdit()
        self.scan_nie_exp_text.setReadOnly(True)
        self.scan_batch_label = QLabel("Batch No : ")
        self.scan_batch_text = QLineEdit()
        self.scan_batch_text.setReadOnly(True)
        self.scan_prod_label = QLabel('Production : ')
        self.scan_prod_text = QLineEdit()
        self.scan_prod_text.setReadOnly(True)
        self.scan_exp_label = QLabel('Expired : ')
        self.scan_exp_text = QLineEdit()
        self.scan_exp_text.setReadOnly(True)
        self.scan_serial_label = QLabel('Serial No : ')
        self.scan_serial_text = QTextEdit()
        self.scan_serial_text.setReadOnly(True)
        self.scan_serial_text.setFixedHeight(100)

    def layouts(self):
        # tab1
        self.mainLayout = QHBoxLayout()
        self.mainLeftLayout = QVBoxLayout()
        self.mainRightLayout = QVBoxLayout()
        self.rightTopLayout = QFormLayout()
        self.rightMiddleLayout = QFormLayout()
        self.topGroupBox = QGroupBox("Generate Codes")
        self.topGroupBox.setContentsMargins(20, 40, 20, 20)
        self.middleGroupBox = QGroupBox("Generate S/N Only")
        self.middleGroupBox.setContentsMargins(20, 40, 20, 20)

        self.mainLeftLayout.addWidget(self.codesTable)
        self.mainRightLayout.addWidget(self.topGroupBox, 60)
        self.mainRightLayout.addWidget(self.middleGroupBox, 40)
        self.mainLayout.addLayout(self.mainLeftLayout, 70)
        self.mainLayout.addLayout(self.mainRightLayout, 30)
        self.tab1.setLayout(self.mainLayout)

        self.rightTopLayout.addRow("NIE:", self.nieEntry)
        self.rightTopLayout.addRow("NIE Expired:", self.nieExpiryEntry)
        self.rightTopLayout.addRow("Batch No:", self.batchEntry)
        self.rightTopLayout.addRow("Production:", self.productionDateEntry)
        self.rightTopLayout.addRow("Expired:", self.expiryDateEntry)
        self.rightTopLayout.addRow("Quantity:", self.codeQuantityEntry)
        self.rightTopLayout.addRow("", self.generateButton)
        self.rightTopLayout.addRow("", self.resetButton)
        self.rightTopLayout.addRow("", self.exportButton)
        self.topGroupBox.setLayout(self.rightTopLayout)

        self.rightMiddleLayout.addRow("Number of Serials:", self.serialQuantityEntry)
        self.rightMiddleLayout.addRow("", self.serialGenExportButton)
        self.middleGroupBox.setLayout(self.rightMiddleLayout)

        # Tab2
        self.scanMainLayout = QHBoxLayout()
        self.scanMainLeftLayout = QVBoxLayout()
        self.scanMainRightLayout = QVBoxLayout()
        self.scanRightTopLayout = QFormLayout()
        self.scanRightMiddleLayout = QFormLayout()
        self.scanTopGroupBox = QGroupBox("Scanned Code")
        self.scanTopGroupBox.setContentsMargins(20, 40, 20, 20)
        self.scanMiddleGroupBox = QGroupBox("Reject Counter")
        self.scanMiddleGroupBox.setContentsMargins(20, 40, 20, 20)

        self.scanMainLeftLayout.addWidget(self.scannedCodesTable)
        self.scanRightTopLayout.addRow(self.scan_nie_label, self.scan_nie_text)
        self.scanRightTopLayout.addRow(self.scan_nie_exp_label, self.scan_nie_exp_text)
        self.scanRightTopLayout.addRow(self.scan_batch_label, self.scan_batch_text)
        self.scanRightTopLayout.addRow(self.scan_prod_label, self.scan_prod_text)
        self.scanRightTopLayout.addRow(self.scan_exp_label, self.scan_exp_text)
        self.scanRightTopLayout.addRow(self.scan_serial_label, self.scan_serial_text)

        self.scanTopGroupBox.setLayout(self.scanRightTopLayout)
        self.scanMainRightLayout.addWidget(self.scanTopGroupBox, 60)
        self.scanMainRightLayout.addWidget(self.scanMiddleGroupBox, 40)
        self.scanMainLayout.addLayout(self.scanMainLeftLayout, 70)
        self.scanMainLayout.addLayout(self.scanMainRightLayout, 30)
        self.tab2.setLayout(self.scanMainLayout)

    def resetCodeGen(self):
        self.nieEntry.setText("")
        self.batchEntry.setText("")
        self.codeQuantityEntry.setText("")
        self.exportButton.setEnabled(False)

        row_count = self.codesTable.rowCount()
        for row in reversed(range(0, row_count)):
            self.codesTable.removeRow(row)

    def generateCodes(self):
        global codes

        nie = self.nieEntry.text()
        nieExpired = self.nieExpiryEntry.text()
        batchNo = self.batchEntry.text()
        prodDate = self.productionDateEntry.text()
        expiredDate = self.expiryDateEntry.text()
        quantity = self.codeQuantityEntry.text()

        codes = []

        if nie and nieExpired and batchNo and quantity != "":
            for i in reversed(range(self.codesTable.rowCount())):
                self.codesTable.removeRow(i)

            for qty in range(0, int(quantity)):
                codes.append((nie, nieExpired, batchNo, prodDate, expiredDate, secrets.token_hex(16)))

            for code in codes:
                row_number = self.codesTable.rowCount()
                self.codesTable.insertRow(row_number)
                for column_number, data in enumerate(code):
                    self.codesTable.setItem(row_number, column_number, QTableWidgetItem(data))

            if len(codes) > 0:
                self.exportButton.setEnabled(True)

        else:
            QMessageBox.information(self, "Warning", "Fields can't be empty")

    def export_to_csv(self):
        global codes
        raw_codes = []
        date_time = QDateTime.currentDateTime().toString("yyMMddhhmmss")

        for code in codes:
            raw_nie_expiry = code[1].split('/')[2] + code[1].split('/')[1] + code[1].split('/')[0]
            raw_prod_date = code[3].split('/')[2] + code[3].split('/')[1] + code[3].split('/')[0]
            raw_expired_date = code[4].split('/')[2] + code[4].split('/')[1] + code[4].split('/')[0]

            raw_codes.append(('(90)'+code[0], '(91)'+raw_nie_expiry, '(10)'+code[2], '(11)'+raw_prod_date, '(17)'+raw_expired_date, '(21)'+code[5]))

        formatted_file_name = 'codes' + date_time + '.csv'
        file_name = QFileDialog.getSaveFileName(self, 'Save File', formatted_file_name, 'CSV (*.csv)')

        if file_name[0] != "":
            try:
                with open(file_name[0], 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerows(raw_codes)

                QMessageBox.information(self, "Success", "CSV file was generated!")
            except Exception as e:
                print(str(e))
        else:
            QMessageBox.information(self, "Information", "Export was cancelled")

    def serial_gen_export(self):
        if self.serialQuantityEntry.text() != "":
            num_serial = int(self.serialQuantityEntry.text())
            serials = []
            for i in range(0, num_serial):
                serials.append((secrets.token_hex(16),))

            date_time = QDateTime.currentDateTime().toString("yyMMddhhmmss")
            formatted_file_name = 'serials' + date_time + '.csv'

            file_name = QFileDialog.getSaveFileName(self, 'Save File', formatted_file_name, 'CSV (*.csv)')
            if file_name[0] != "":
                try:
                    with open(file_name[0], 'w', newline='') as f:
                        writer = csv.writer(f)
                        writer.writerows(serials)

                    QMessageBox.information(self, 'Success', 'CSV has been exported')
                except:
                    QMessageBox.information(self, 'Error', 'CSV has not been exported')

        else:
            QMessageBox.information(self, "Warning", "Quantity field can't be empty!")

    def serial_quantity_changed(self):
        if self.serialQuantityEntry.text() != "":
            self.serialGenExportButton.setEnabled(True)
        else:
            self.serialGenExportButton.setEnabled(False)

    def setScanner(self):
        self.newScanner = scanner.SetScanner()
        self.newScanner.connect_signal.connect(self.connect_scanner)
        self.newScanner.disconnect_signal.connect(self.disconnect_scanner)

    def AgScanner(self):
        self.newAgScanner = ag_scanner.SetAgScanner()

    # @staticmethod
    def connect_scanner(self, port):
        self.scanner = QSerialPort(port, baudRate=QSerialPort.Baud115200, readyRead=self.scanReceive)
        if not self.scanner.isOpen():
            res = self.scanner.open(QIODevice.ReadWrite)
            config.scanner_connected = True
            if not res:
                print('Open port gagal')

    # @staticmethod
    def disconnect_scanner(self):
        self.scanner.close()
        config.scanner_connected = False

    def scanReceive(self):
        while self.scanner.canReadLine():
            self.code_text = self.scanner.readLine().data().decode()
            self.code_text = self.code_text.rstrip('\r\n')

        try:
            self.get_code(self.code_text)
            print(self.code)
        except Exception as e:
            print(str(e))

        try:
            row_number = self.scannedCodesTable.rowCount()
            self.scannedCodesTable.insertRow(row_number)
            for column_number, data in enumerate(self.code):
                self.scannedCodesTable.setItem(row_number, column_number, QTableWidgetItem(data))
        except Exception as e:
            print(str(e))

    def get_code(self, text):
        nie_pattern = '\(90\)\w*'
        nie_exp_pattern = '\(91\)\w*'
        batch_pattern = '\(10\)\w*'
        prod_pattern = '\(11\)\w*'
        exp_pattern = '\(17\)\w*'
        serial_pattern = '\(21\)\w*'

        self.code = []

        self.nie = self.get_sub_code(nie_pattern, text)
        self.nie_exp = self.get_sub_code(nie_exp_pattern, text)
        self.batch = self.get_sub_code(batch_pattern, text)
        self.prod = self.get_sub_code(prod_pattern, text)
        self.exp = self.get_sub_code(exp_pattern, text)
        self.serial = self.get_sub_code(serial_pattern, text)

        self.code.append(self.nie)
        self.code.append(self.nie_exp)
        self.code.append(self.batch)
        self.code.append(self.prod)
        self.code.append(self.exp)
        self.code.append(self.serial)

    def get_sub_code(self, pattern, text):
        res = re.findall(pattern, text)
        return res[0][4:]


def main():
    App = QApplication(sys.argv)
    window = Main()
    sys.exit(App.exec_())


if __name__ == "__main__":
    main()

# Happy New Year 2020 - Working at Midnight :-)
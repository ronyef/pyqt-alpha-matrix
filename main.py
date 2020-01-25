import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtSerialPort import *
import secrets
import csv
import scanner
import ag_scanner
import rejector
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
        self.setAgScannerAction = QAction(QIcon('icons/qr-code-scan-24'), "Set Aggregate", self)
        self.setAgScannerAction.triggered.connect(self.setAgScanner)
        self.toolbar.addAction(self.setAgScannerAction)
        self.setRejectorAction = QAction(QIcon('icons/trash.png'), "Set Rejector", self)
        self.setRejectorAction.triggered.connect(self.setRejector)
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
        self.codesTable.setEditTriggers(QTableWidget.NoEditTriggers)
        self.codesTable.setHorizontalHeaderItem(0, QTableWidgetItem("NIE"))
        self.codesTable.setHorizontalHeaderItem(1, QTableWidgetItem("NIE Exp"))
        self.codesTable.setHorizontalHeaderItem(2, QTableWidgetItem("Batch"))
        self.codesTable.setHorizontalHeaderItem(3, QTableWidgetItem("Prod Date"))
        self.codesTable.setHorizontalHeaderItem(4, QTableWidgetItem("Expired"))
        self.codesTable.setHorizontalHeaderItem(5, QTableWidgetItem("Serial"))
        self.codesTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.codesTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.codesTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.codesTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.codesTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.codesTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)

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
        self.resetButton.setEnabled(False)
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
        self.scannedCodesTable.setEditTriggers(QTableWidget.NoEditTriggers)
        self.scannedCodesTable.setHorizontalHeaderItem(0, QTableWidgetItem("NIE"))
        self.scannedCodesTable.setHorizontalHeaderItem(1, QTableWidgetItem("NIE Exp"))
        self.scannedCodesTable.setHorizontalHeaderItem(2, QTableWidgetItem("Batch"))
        self.scannedCodesTable.setHorizontalHeaderItem(3, QTableWidgetItem("Prod Date"))
        self.scannedCodesTable.setHorizontalHeaderItem(4, QTableWidgetItem("Expired"))
        self.scannedCodesTable.setHorizontalHeaderItem(5, QTableWidgetItem("Serial"))
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.scannedCodesTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)

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
        self.scanExportButton = QPushButton("Export CSV")
        self.scanExportButton.clicked.connect(self.scanExport)
        self.scanExportButton.setEnabled(False)
        self.scanResetButton = QPushButton("Reset")
        self.scanResetButton.setEnabled(False)
        self.scanResetButton.clicked.connect(self.resetScanTable)

        # Scan bottom widgets
        self.scanRejectTitle = QLabel("Reject :")
        self.scan_reject_lcd = QLCDNumber()
        self.scanSuccessTitle = QLabel("Scanned :")
        self.scan_success_lcd = QLCDNumber()
        self.scanRateTitle = QLabel("Accuracy :")
        self.scan_rate_lcd = QLCDNumber()
        self.scanPercentLabel = QLabel("%")
        self.scanCounterResetButton = QPushButton("Reset Counter")
        self.scanCounterResetButton.clicked.connect(self.resetScanCounter)

        # AGGREGATION WIDGET
        self.aggregationCodesTable = QTableWidget()
        self.aggregationCodesTable.setColumnCount(6)
        self.aggregationCodesTable.setEditTriggers(QTableWidget.NoEditTriggers)
        self.aggregationCodesTable.setHorizontalHeaderItem(0, QTableWidgetItem("NIE"))
        self.aggregationCodesTable.setHorizontalHeaderItem(1, QTableWidgetItem("NIE Exp"))
        self.aggregationCodesTable.setHorizontalHeaderItem(2, QTableWidgetItem("Batch"))
        self.aggregationCodesTable.setHorizontalHeaderItem(3, QTableWidgetItem("Prod Date"))
        self.aggregationCodesTable.setHorizontalHeaderItem(4, QTableWidgetItem("Expired"))
        self.aggregationCodesTable.setHorizontalHeaderItem(5, QTableWidgetItem("Serial"))
        self.aggregationCodesTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.aggregationCodesTable.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.aggregationCodesTable.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.aggregationCodesTable.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.aggregationCodesTable.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.aggregationCodesTable.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)

        # Scan top right widgets
        self.aggregate_nie_label = QLabel("NIE : ")
        self.aggregate_nie_text = QLineEdit()
        self.aggregate_nie_text.setReadOnly(True)
        self.aggregate_nie_exp_label = QLabel('NIE Expired : ')
        self.aggregate_nie_exp_text = QLineEdit()
        self.aggregate_nie_exp_text.setReadOnly(True)
        self.aggregate_batch_label = QLabel("Batch No : ")
        self.aggregate_batch_text = QLineEdit()
        self.aggregate_batch_text.setReadOnly(True)
        self.aggregate_prod_label = QLabel('Production : ')
        self.aggregate_prod_text = QLineEdit()
        self.aggregate_prod_text.setReadOnly(True)
        self.aggregate_exp_label = QLabel('Expired : ')
        self.aggregate_exp_text = QLineEdit()
        self.aggregate_exp_text.setReadOnly(True)
        self.aggregate_serial_label = QLabel('Serial No : ')
        self.aggregate_serial_text = QTextEdit()
        self.aggregate_serial_text.setReadOnly(True)
        self.aggregate_serial_text.setFixedHeight(100)
        self.aggregate_print_button = QPushButton("Print Label")
        self.aggregate_print_button.setEnabled(False)
        self.aggregate_export_button = QPushButton("Export CSV")
        self.aggregate_export_button.setEnabled(False)
        self.aggregate_reset_button = QPushButton("Reset")
        self.aggregate_reset_button.setEnabled(False)
        self.aggregate_reset_button.clicked.connect(self.reset_aggregation_table)

        # Aggregation bottom widgets
        self.aggregate_count_lcd = QLCDNumber()
        self.aggregate_count_lcd.setMinimumHeight(80)
        self.aggregate_counter_reset_button = QPushButton("Reset")
        self.aggregate_counter_reset_button.clicked.connect(self.reset_aggregate_counter)

        # Statistics widgets
        self.statisticsGenCodesLCD = QLCDNumber()
        self.statisticsGenCodesLCD.setStyleSheet("QLCDNumber {background-color: yellow}")
        self.statisticsGenSerialsLCD = QLCDNumber()
        self.statisticsGenSerialsLCD.setStyleSheet("QLCDNumber {background-color: yellow}")
        self.statisticsScannedLCD = QLCDNumber()
        self.statisticsScannedLCD.setStyleSheet("QLCDNumber {background-color: green; color: white;}")
        self.statisticsRejectLCD = QLCDNumber()
        self.statisticsRejectLCD.setStyleSheet("QLCDNumber {background-color: red; color: white;}")
        self.statisticsChildLCD = QLCDNumber()
        self.statisticsChildLCD.setStyleSheet("QLCDNumber {background-color: yellow}")
        self.statisticsParentLCD = QLCDNumber()
        self.statisticsParentLCD.setStyleSheet("QLCDNumber {background-color: yellow}")
        self.statisticsGenCodesLabel = QLabel("Generated Codes")
        self.statisticsGenCodesLabel.setAlignment(Qt.AlignCenter)
        self.statisticsGenSerialsLabel = QLabel("Generated Serials")
        self.statisticsGenSerialsLabel.setAlignment(Qt.AlignCenter)
        self.statisticsScannedLabel = QLabel("Scanned Codes")
        self.statisticsScannedLabel.setAlignment(Qt.AlignCenter)
        self.statisticsRejectLabel = QLabel("Rejected Codes")
        self.statisticsRejectLabel.setAlignment(Qt.AlignCenter)
        self.statisticsParentLabel = QLabel("Parents")
        self.statisticsParentLabel.setAlignment(Qt.AlignCenter)
        self.statisticsChildLabel = QLabel("Children")
        self.statisticsChildLabel.setAlignment(Qt.AlignCenter)

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
        self.scanRightMiddleLayout = QVBoxLayout()
        self.scanSuccessLayout = QHBoxLayout()
        self.scanRejectLayout = QHBoxLayout()
        self.scanRateLayout = QHBoxLayout()
        self.scanButtonBarLayout = QHBoxLayout()
        self.scanTopGroupBox = QGroupBox("Scanned Code")
        self.scanTopGroupBox.setContentsMargins(20, 40, 20, 20)
        self.scanMiddleGroupBox = QGroupBox("Product Counter")
        self.scanMiddleGroupBox.setContentsMargins(10, 20, 10, 10)

        self.scanMainLeftLayout.addWidget(self.scannedCodesTable)
        self.scanRightTopLayout.addRow(self.scan_nie_label, self.scan_nie_text)
        self.scanRightTopLayout.addRow(self.scan_nie_exp_label, self.scan_nie_exp_text)
        self.scanRightTopLayout.addRow(self.scan_batch_label, self.scan_batch_text)
        self.scanRightTopLayout.addRow(self.scan_prod_label, self.scan_prod_text)
        self.scanRightTopLayout.addRow(self.scan_exp_label, self.scan_exp_text)
        self.scanRightTopLayout.addRow(self.scan_serial_label, self.scan_serial_text)
        self.scanRightTopLayout.addRow('', self.scanResetButton)
        self.scanRightTopLayout.addRow('', self.scanExportButton)

        # Right bottom layout
        self.scanSuccessLayout.addWidget(self.scanSuccessTitle)
        self.scanSuccessLayout.addWidget(self.scan_success_lcd)
        self.scanRejectLayout.addWidget(self.scanRejectTitle)
        self.scanRejectLayout.addWidget(self.scan_reject_lcd)
        self.scanRateLayout.addWidget(self.scanRateTitle, 60)
        self.scanRateLayout.addWidget(self.scan_rate_lcd, 30)
        self.scanRateLayout.addWidget(self.scanPercentLabel, 10)
        self.scanRightMiddleLayout.addLayout(self.scanSuccessLayout)
        self.scanRightMiddleLayout.addLayout(self.scanRejectLayout)
        self.scanRightMiddleLayout.addLayout(self.scanRateLayout)
        self.scanRightMiddleLayout.addWidget(self.scanCounterResetButton)
        self.scanRightMiddleLayout.addLayout(self.scanButtonBarLayout)

        self.scanTopGroupBox.setLayout(self.scanRightTopLayout)
        self.scanMiddleGroupBox.setLayout(self.scanRightMiddleLayout)
        self.scanMainRightLayout.addWidget(self.scanTopGroupBox, 60)
        self.scanMainRightLayout.addWidget(self.scanMiddleGroupBox, 40)
        self.scanMainLayout.addLayout(self.scanMainLeftLayout, 70)
        self.scanMainLayout.addLayout(self.scanMainRightLayout, 30)
        self.tab2.setLayout(self.scanMainLayout)

        # Tab3
        self.aggregateMainLayout = QHBoxLayout()
        self.aggregateMainLeftLayout = QVBoxLayout()
        self.aggregateMainRightLayout = QVBoxLayout()
        self.aggregateRightTopLayout = QFormLayout()
        self.aggregateRightMiddleLayout = QVBoxLayout()
        self.aggregateTopGroupBox = QGroupBox("Scanned Code")
        self.aggregateTopGroupBox.setContentsMargins(20, 40, 20, 20)
        self.aggregateMiddleGroupBox = QGroupBox("Product Counter")
        self.aggregateMiddleGroupBox.setContentsMargins(20, 40, 20, 20)

        self.aggregateMainLeftLayout.addWidget(self.aggregationCodesTable)
        self.aggregateRightTopLayout.addRow(self.aggregate_nie_label, self.aggregate_nie_text)
        self.aggregateRightTopLayout.addRow(self.aggregate_nie_exp_label, self.aggregate_nie_exp_text)
        self.aggregateRightTopLayout.addRow(self.aggregate_batch_label, self.aggregate_batch_text)
        self.aggregateRightTopLayout.addRow(self.aggregate_prod_label, self.aggregate_prod_text)
        self.aggregateRightTopLayout.addRow(self.aggregate_exp_label, self.aggregate_exp_text)
        self.aggregateRightTopLayout.addRow(self.aggregate_serial_label, self.aggregate_serial_text)
        self.aggregateRightTopLayout.addRow('', self.aggregate_reset_button)
        self.aggregateRightTopLayout.addRow('', self.aggregate_print_button)
        self.aggregateRightTopLayout.addRow('', self.aggregate_export_button)

        # Right bottom layout
        self.aggregateRightMiddleLayout.addWidget(self.aggregate_count_lcd)
        self.aggregateRightMiddleLayout.addWidget(self.aggregate_counter_reset_button)

        self.aggregateTopGroupBox.setLayout(self.aggregateRightTopLayout)
        self.aggregateMiddleGroupBox.setLayout(self.aggregateRightMiddleLayout)
        self.aggregateMainRightLayout.addWidget(self.aggregateTopGroupBox, 60)
        self.aggregateMainRightLayout.addWidget(self.aggregateMiddleGroupBox, 40)
        self.aggregateMainLayout.addLayout(self.aggregateMainLeftLayout, 70)
        self.aggregateMainLayout.addLayout(self.aggregateMainRightLayout, 30)
        self.tab3.setLayout(self.aggregateMainLayout)

        # Tab4
        self.statisticsMainLayout = QHBoxLayout()
        self.statisticsLeftLayout = QVBoxLayout()
        self.statisticsRightLayout = QVBoxLayout()
        self.statisticsPrintLayout = QHBoxLayout()
        self.statisticsScanLayout = QHBoxLayout()
        self.statisticsAggregateLayout = QHBoxLayout()
        self.statisticsScanCodeLayout = QVBoxLayout()
        self.statisticsScanRejectLayout = QVBoxLayout()
        self.statisticsPrintCodeLayout = QVBoxLayout()
        self.statisticsPrintSerialsLayout = QVBoxLayout()
        self.statisticsAggregateChildLayout = QVBoxLayout()
        self.statisticsAggregateParentLayout = QVBoxLayout()

        # Add widgets
        self.statisticsPrintCodeLayout.addWidget(self.statisticsGenCodesLCD)
        self.statisticsPrintCodeLayout.addWidget(self.statisticsGenCodesLabel)
        self.statisticsPrintSerialsLayout.addWidget(self.statisticsGenSerialsLCD)
        self.statisticsPrintSerialsLayout.addWidget(self.statisticsGenSerialsLabel)
        self.statisticsScanCodeLayout.addWidget(self.statisticsScannedLCD)
        self.statisticsScanCodeLayout.addWidget(self.statisticsScannedLabel)
        self.statisticsScanRejectLayout.addWidget(self.statisticsRejectLCD)
        self.statisticsScanRejectLayout.addWidget(self.statisticsRejectLabel)
        self.statisticsAggregateParentLayout.addWidget(self.statisticsParentLCD)
        self.statisticsAggregateParentLayout.addWidget(self.statisticsParentLabel)
        self.statisticsAggregateChildLayout.addWidget(self.statisticsChildLCD)
        self.statisticsAggregateChildLayout.addWidget(self.statisticsChildLabel)

        self.statisticsPrintLayout.addLayout(self.statisticsPrintCodeLayout)
        self.statisticsPrintLayout.addLayout(self.statisticsPrintSerialsLayout)
        self.statisticsScanLayout.addLayout(self.statisticsScanCodeLayout)
        self.statisticsScanLayout.addLayout(self.statisticsScanRejectLayout)
        self.statisticsAggregateLayout.addLayout(self.statisticsAggregateParentLayout)
        self.statisticsAggregateLayout.addLayout(self.statisticsAggregateChildLayout)

        self.statisticsPrintGroup = QGroupBox("Print")
        self.statisticsPrintGroup.setContentsMargins(20, 40, 20, 20)
        self.statisticsScanGroup = QGroupBox("Scan")
        self.statisticsScanGroup.setContentsMargins(20, 40, 20, 20)
        self.statisticsAggregateGroup = QGroupBox("Aggregation")
        self.statisticsAggregateGroup.setContentsMargins(20, 40, 20, 20)

        self.statisticsPrintGroup.setLayout(self.statisticsPrintLayout)
        self.statisticsScanGroup.setLayout(self.statisticsScanLayout)
        self.statisticsAggregateGroup.setLayout(self.statisticsAggregateLayout)

        self.statisticsLeftLayout.addWidget(self.statisticsPrintGroup)
        self.statisticsLeftLayout.addWidget(self.statisticsScanGroup)
        self.statisticsLeftLayout.addWidget(self.statisticsAggregateGroup)
        self.statisticsMainLayout.addLayout(self.statisticsLeftLayout, 50)
        self.statisticsMainLayout.addLayout(self.statisticsRightLayout, 50)
        self.tab4.setLayout(self.statisticsMainLayout)

    def resetCodeGen(self):
        self.nieEntry.setText("")
        self.batchEntry.setText("")
        self.codeQuantityEntry.setText("")
        self.exportButton.setEnabled(False)
        self.resetButton.setEnabled(False)

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
                self.resetButton.setEnabled(True)

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

            raw_codes.append(('(90)' + code[0], '(91)' + raw_nie_expiry, '(10)' + code[2], '(11)' + raw_prod_date,
                              '(17)' + raw_expired_date, '(21)' + code[5]))

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

    def setAgScanner(self):
        self.newAgScanner = ag_scanner.AgScanner()
        self.newAgScanner.connect_signal.connect(self.connect_ag_scanner)
        self.newAgScanner.disconnect_signal.connect(self.disconnect_ag_scanner)

    def connect_scanner(self, port):
        if config.ag_scanner_port != port:
            self.scanner = QSerialPort(port, baudRate=QSerialPort.Baud115200, readyRead=self.scanReceive)
            if not self.scanner.isOpen():
                res = self.scanner.open(QIODevice.ReadWrite)
                config.scanner_connected = True
                config.scanner_port = port
                if not res:
                    print('Open port gagal')
        else:
            # self.newScanner.connectButton.setChecked(False)
            self.newScanner.close()
            QMessageBox.information(self, "Serial Port", "Port has already been used!")

    def disconnect_scanner(self):
        self.scanner.close()
        config.scanner_connected = False
        config.scanner_port = ''

    def scanReceive(self):
        line_code = []

        while self.scanner.canReadLine():
            self.code_text = self.scanner.readLine().data().decode()
            self.code_text = self.code_text.rstrip('\r\n')

        print(self.code_text)

        if self.code_text == 'ERROR':
            if config.scan_count > 0:
                config.scan_reject += 1
                self.update_scan_lcd()
                return

            return


        line_code = self.get_line_code(self.code_text)

        # Insert data to table on tab2
        row_number = self.scannedCodesTable.rowCount()
        self.scannedCodesTable.insertRow(row_number)
        for column_number, data in enumerate(line_code):
            self.scannedCodesTable.setItem(row_number, column_number, QTableWidgetItem(data))

        self.scan_nie_text.setText(line_code[0])
        self.scan_nie_exp_text.setText(line_code[1])
        self.scan_batch_text.setText(line_code[2])
        self.scan_prod_text.setText(line_code[3])
        self.scan_exp_text.setText(line_code[4])
        self.scan_serial_text.setText(line_code[5])

        # Enable reset and export buttons if otherwise
        if not self.scanResetButton.isEnabled():
            self.scanResetButton.setEnabled(True)

        if not self.scanExportButton.isEnabled():
            self.scanExportButton.setEnabled(True)

        # Update counter
        config.scan_count += 1
        self.update_scan_lcd()

    def update_scan_lcd(self):
        try:
            config.scan_rate = round((config.scan_count - config.scan_reject) / config.scan_count * 100, 2)
            scanned = str(config.scan_count)
            scan_rate = str(config.scan_rate)
            scan_reject = str(config.scan_reject)

            self.scan_success_lcd.display(scanned)
            self.scan_reject_lcd.display(scan_reject)
            self.scan_rate_lcd.display(scan_rate)
        except Exception as e:
            print('Terjadi kesalahan: ' + str(e))

    def get_sub_code(self, pattern, text):
        res = re.findall(pattern, text)
        return res[0][4:]

    def resetScanCounter(self):
        config.scan_reject = 0
        reject = str(config.scan_reject)
        count = str(config.scan_count)
        rate = str(config.scan_rate)
        config.scan_count = 0
        config.scan_rate = 0
        self.scan_success_lcd.display(count)
        self.scan_reject_lcd.display(reject)
        self.scan_rate_lcd.display(rate)

    def resetScanTable(self):
        confirm = QMessageBox.question(self, "Warning", "Are you sure to reset scanned codes?",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm == QMessageBox.Yes:
            row_count = self.scannedCodesTable.rowCount()
            for row in reversed(range(0, row_count)):
                self.scannedCodesTable.removeRow(row)

            self.scan_nie_text.setText('')
            self.scan_nie_exp_text.setText('')
            self.scan_batch_text.setText('')
            self.scan_prod_text.setText('')
            self.scan_exp_text.setText('')
            self.scan_serial_text.setText('')

            self.scanResetButton.setEnabled(False)
            self.scanExportButton.setEnabled(False)

    def scanExport(self):
        date_time = QDateTime.currentDateTime().toString("yyMMddhhmmss")
        formatted_file_name = 'scanned' + date_time + '.csv'
        file_name = QFileDialog.getSaveFileName(self, 'Save File', formatted_file_name, 'CSV (*.csv)')

        if file_name[0] != "":
            row_num = self.scannedCodesTable.rowCount()
            codes = []

            for row in range(0, row_num):
                code_line = []

                for col in range(0, 6):
                    code_line.append(self.scannedCodesTable.item(row, col).text())

                codes.append(code_line)

            try:
                with open(file_name[0], 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerows(codes)

                QMessageBox.information(self, 'Success', 'CSV has been exported')
            except:
                QMessageBox.information(self, 'Error', 'CSV has not been exported')

    def connect_ag_scanner(self, port):
        if config.scanner_port != port:
            self.ag_scanner = QSerialPort(port, baudRate=QSerialPort.Baud115200, readyRead=self.agScanReceive)
            if not self.ag_scanner.isOpen():
                res = self.ag_scanner.open(QIODevice.ReadWrite)
                config.ag_scanner_connected = True
                config.ag_scanner_port = port
                if not res:
                    print('Open port gagal')
        else:
            self.newAgScanner.close()
            QMessageBox.information(self, "Serial Port", "Port has been used")

    def disconnect_ag_scanner(self):
        self.ag_scanner.close()
        config.ag_scanner_connected = False
        config.ag_scanner_port = ''

    def agScanReceive(self):
        line_code = []

        while self.ag_scanner.canReadLine():
            self.ag_code_text = self.ag_scanner.readLine().data().decode()
            self.ag_code_text = self.ag_code_text.rstrip('\r\n')

        line_code = self.get_line_code(self.ag_code_text)

        self.aggregate_nie_text.setText(line_code[0])
        self.aggregate_nie_exp_text.setText(line_code[1])
        self.aggregate_batch_text.setText(line_code[2])
        self.aggregate_prod_text.setText(line_code[3])
        self.aggregate_exp_text.setText(line_code[4])
        self.aggregate_serial_text.setText(line_code[5])

        # Insert data to table on tab3
        row_number = self.aggregationCodesTable.rowCount()
        self.aggregationCodesTable.insertRow(row_number)
        for column_number, data in enumerate(line_code):
            self.aggregationCodesTable.setItem(row_number, column_number, QTableWidgetItem(data))

        # Enable reset & export buttons if otherwise
        if not self.aggregate_reset_button.isEnabled():
            self.aggregate_reset_button.setEnabled(True)

        if not self.aggregate_export_button.isEnabled():
            self.aggregate_export_button.setEnabled(True)

        if not self.aggregate_print_button.isEnabled():
            self.aggregate_print_button.setEnabled(True)

        # Update counter
        config.ag_scan_count += 1
        scanned = str(config.ag_scan_count)

        self.aggregate_count_lcd.display(scanned)

    def get_line_code(self, text):
        nie_pattern = '\(90\)\w*'
        nie_exp_pattern = '\(91\)\w*'
        batch_pattern = '\(10\)\w*'
        prod_pattern = '\(11\)\w*'
        exp_pattern = '\(17\)\w*'
        serial_pattern = '\(21\)\w*'

        line_code = []

        nie = self.get_sub_code(nie_pattern, text)
        nie_exp = self.get_sub_code(nie_exp_pattern, text)
        nie_exp = nie_exp[6:] + '/' + nie_exp[4:6] + '/' + nie_exp[:4]
        batch = self.get_sub_code(batch_pattern, text)
        prod = self.get_sub_code(prod_pattern, text)
        prod = prod[6:] + '/' + prod[4:6] + '/' + prod[:4]
        exp = self.get_sub_code(exp_pattern, text)
        exp = exp[6:] + '/' + exp[4:6] + '/' + exp[:4]
        serial = self.get_sub_code(serial_pattern, text)

        line_code.append(nie)
        line_code.append(nie_exp)
        line_code.append(batch)
        line_code.append(prod)
        line_code.append(exp)
        line_code.append(serial)

        return line_code

    def reset_aggregation_table(self):
        confirm = QMessageBox.question(self, "Warning", "Are you sure to reset table?",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if confirm == QMessageBox.Yes:
            row_count = self.aggregationCodesTable.rowCount()
            for row in reversed(range(0, row_count)):
                self.aggregationCodesTable.removeRow(row)

            self.aggregate_nie_text.setText('')
            self.aggregate_nie_exp_text.setText('')
            self.aggregate_batch_text.setText('')
            self.aggregate_prod_text.setText('')
            self.aggregate_exp_text.setText('')
            self.aggregate_serial_text.setText('')

            self.aggregate_reset_button.setEnabled(False)
            self.aggregate_export_button.setEnabled(False)
            self.aggregate_print_button.setEnabled(False)

    def reset_aggregate_counter(self):
        config.ag_scan_count = 0
        count = str(config.ag_scan_count)
        self.aggregate_count_lcd.display(count)

    def setRejector(self):
        self.newRejector = rejector.Rejector()
        self.newRejector.connect_signal.connect(self.connect_rejector)
        self.newRejector.disconnect_signal.connect(self.disconnect_rejector)

    def connect_rejector(self, port):
        if config.rejector_port != port:
            self.rejector = QSerialPort(port, baudRate=QSerialPort.Baud115200, readyRead=self.rejectorReceive)
            if not self.rejector.isOpen():
                res = self.rejector.open(QIODevice.ReadWrite)
                config.rejector_connected = True
                config.rejector_port = port
                if not res:
                    print('Open port gagal')
        else:
            self.newAgScanner.close()
            QMessageBox.information(self, "Serial Port", "Port has been used")

    def disconnect_rejector(self):
        self.rejector.close()
        config.rejector_connected = False
        config.rejector_port = ''

    def rejectorReceive(self):
        pass


def main():
    App = QApplication(sys.argv)
    window = Main()
    sys.exit(App.exec_())


if __name__ == "__main__":
    main()

# Happy New Year 2020 - Working at Midnight :-)

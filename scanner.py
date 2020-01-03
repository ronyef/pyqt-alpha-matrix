import sys
from PyQt5.Qt import *
from PyQt5.QtWidgets import *
import serial.tools.list_ports
import style
import config


class SetScanner(QWidget):
    connect_signal = pyqtSignal(str)
    disconnect_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Scanner")
        self.setWindowIcon(QIcon('icons/barcode-scanner-32.png'))
        self.setGeometry(460, 200, 350, 550)
        self.setFixedSize(self.size())
        self.UI()
        self.show()

    def serial_port(self):
        self.ports = serial.tools.list_ports.comports(include_links=False)
        # self.availablePorts = []
        config.ports = []
        for port in self.ports:
            config.ports.append(port.device)

        # config.ports = self.availablePorts
        self.portCombo.addItems(config.ports)

    def UI(self):
        self.widgets()
        self.layouts()
        self.serial_port()

    def widgets(self):
        self.titleText = QLabel("Set Scanner")
        self.titleText.setAlignment(Qt.AlignCenter)
        self.scannerImage = QLabel()
        scan_img = QPixmap('images/scanner1.jpg')
        self.scannerImage.setPixmap(scan_img)
        self.scannerImage.setAlignment(Qt.AlignCenter)

        # Bottom Widgets
        self.portText = QLabel("Select Port: ")
        self.portCombo = QComboBox()
        self.connectButton = QPushButton("Connect")
        self.connectButton.setCheckable(True)
        self.connectButton.toggled.connect(self.on_toggled)

        if config.scanner_connected:
            self.connectButton.setChecked(True)
        else:
            self.connectButton.setChecked(False)

    def layouts(self):
        self.mainLayout = QVBoxLayout()
        self.topLayout = QVBoxLayout()
        self.bottomLayout = QFormLayout()
        self.topFrame = QFrame()
        self.topFrame.setStyleSheet(style.setDeviceTopFrame())
        self.topFrame.setContentsMargins(20, 40, 20, 20)
        self.bottomFrame = QFrame()
        self.bottomFrame.setStyleSheet(style.setDeviceBottomFrame())
        self.bottomFrame.setContentsMargins(20, 40, 20, 20)

        # Top layout widgets
        self.topLayout.addWidget(self.titleText)
        self.topLayout.addWidget(self.scannerImage)

        # Bottom layout widgets
        self.bottomLayout.addRow(self.portText, self.portCombo)
        self.bottomLayout.addRow('', self.connectButton)

        self.topFrame.setLayout(self.topLayout)
        self.bottomFrame.setLayout(self.bottomLayout)
        self.mainLayout.addWidget(self.topFrame)
        self.mainLayout.addWidget(self.bottomFrame)
        self.setLayout(self.mainLayout)

    # @QtCore.pyqtSlot(bool)
    def on_toggled(self, checked):
        self.connectButton.setText('Disconnect' if checked else 'Connect')

        if checked:
            self.connect_signal.emit(self.portCombo.currentText())
        else:
            self.disconnect_signal.emit()


def main():
    App = QApplication(sys.argv)
    window = SetScanner()
    sys.exit(App.exec_())


if __name__ == '__main__':
    main()

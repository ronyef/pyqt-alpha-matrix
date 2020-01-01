import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.Qt import *
import style
import config
import serial.tools.list_ports


class SetAgScanner(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Scanner")
        self.setWindowIcon(QIcon('icons/barcode-scanner-32.png'))
        self.setGeometry(460, 200, 350, 550)
        self.setFixedSize(self.size())
        self.UI()
        self.serial_port()
        self.show()

    def serial_port(self):
        self.ports = serial.tools.list_ports.comports(include_links=False)
        for port in self.ports:
            config.ports.append(port.device)

        self.portCombo.addItems(config.ports)

    def UI(self):
        self.widgets()
        self.layouts()

    def widgets(self):
        self.titleText = QLabel("Set Scanner")
        self.titleText.setAlignment(Qt.AlignCenter)
        self.scannerImage = QLabel()
        scanImg = QPixmap('images/portable_scanner.jpg')
        self.scannerImage.setPixmap(scanImg)
        self.scannerImage.setAlignment(Qt.AlignCenter)

        # Bottom Widgets
        self.portText = QLabel("Select Port: ")
        self.portCombo = QComboBox()
        self.connectButton = QPushButton("Connect")
        self.disconnectButton = QPushButton("Disconnect")

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
        self.bottomLayout.addRow('', self.disconnectButton)

        self.topFrame.setLayout(self.topLayout)
        self.bottomFrame.setLayout(self.bottomLayout)
        self.mainLayout.addWidget(self.topFrame)
        self.mainLayout.addWidget(self.bottomFrame)
        self.setLayout(self.mainLayout)


def main():
    App = QApplication(sys.argv)
    window = Window()
    sys.exit(App.exec_())


if __name__ == '__main__':
    main()
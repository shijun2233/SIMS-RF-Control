import sys
import serial
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QTextEdit, QLineEdit, QComboBox, QLabel, QMessageBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal


class SerialThread(QThread):
    response_received = pyqtSignal(str)

    def __init__(self, port, baudrate):
        super().__init__()
        self.port = port
        self.baudrate = baudrate
        self.serial_conn = None
        self.running = True

    def run(self):
        try:
            self.serial_conn = serial.Serial(self.port, self.baudrate, timeout=2)
            while self.running:
                if self.serial_conn.in_waiting > 0:
                    response = self.serial_conn.readline().decode('utf-8').strip()
                    self.response_received.emit(response)
        except Exception as e:
            self.response_received.emit(f"[ERROR] {str(e)}")

    def send_command(self, command):
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.write((command + '\n').encode('utf-8'))

    def stop(self):
        self.running = False
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
        self.quit()
        self.wait()


class SCPIGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Keithley 6485/6487 SCPI 串口测试工具")
        self.setGeometry(100, 100, 800, 600)
        self.thread = None

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 串口配置
        serial_layout = QHBoxLayout()
        serial_layout.addWidget(QLabel("串口端口:"))
        self.port_combo = QComboBox()
        self.port_combo.addItems(["COM3", "COM4", "COM5", "COM6"])  # 可手动添加更多
        serial_layout.addWidget(self.port_combo)

        serial_layout.addWidget(QLabel("波特率:"))
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(["9600", "19200", "38400", "115200"])
        self.baud_combo.setCurrentText("9600")
        serial_layout.addWidget(self.baud_combo)

        self.connect_btn = QPushButton("连接")
        self.connect_btn.clicked.connect(self.toggle_connection)
        serial_layout.addWidget(self.connect_btn)

        layout.addLayout(serial_layout)

        # 命令输入
        self.command_input = QLineEdit()
        self.command_input.setPlaceholderText("输入 SCPI 命令，例如: *IDN?")
        layout.addWidget(self.command_input)

        # 发送按钮
        self.send_btn = QPushButton("发送命令")
        self.send_btn.clicked.connect(self.send_command)
        layout.addWidget(self.send_btn)

        # 响应显示
        self.response_log = QTextEdit()
        self.response_log.setReadOnly(True)
        layout.addWidget(self.response_log)

        # 清除按钮
        self.clear_btn = QPushButton("清除日志")
        self.clear_btn.clicked.connect(self.response_log.clear)
        layout.addWidget(self.clear_btn)

        self.setLayout(layout)

    def toggle_connection(self):
        if self.thread is None or not self.thread.isRunning():
            port = self.port_combo.currentText()
            baud = int(self.baud_combo.currentText())
            self.thread = SerialThread(port, baud)
            self.thread.response_received.connect(self.display_response)
            self.thread.start()
            self.connect_btn.setText("断开")
            self.response_log.append(f"[INFO] 已连接到 {port} @ {baud} baud")
        else:
            self.thread.stop()
            self.connect_btn.setText("连接")
            self.response_log.append("[INFO] 串口已断开")

    def send_command(self):
        if self.thread and self.thread.isRunning():
            command = self.command_input.text().strip()
            if command:
                self.thread.send_command(command)
                self.response_log.append(f"[发送] {command}")
                self.command_input.clear()
        else:
            QMessageBox.warning(self, "警告", "串口未连接，请先连接设备。")

    def display_response(self, response):
        self.response_log.append(f"[接收] {response}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SCPIGUI()
    window.show()
    sys.exit(app.exec_())
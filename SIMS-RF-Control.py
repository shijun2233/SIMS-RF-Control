import sys
import os
import time
import math
import random
import logging
import threading
from typing import Optional, Union
from collections import deque
from datetime import datetime
import time
import logging
import threading
import sys
import os

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QKeyEvent
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QGridLayout, QGroupBox, QDoubleSpinBox, QPushButton, QTextEdit,
    QFormLayout, QMainWindow, QFileDialog, QAction, QMenu, QLineEdit,
    QComboBox, QDialog, QDialogButtonBox, QMessageBox, QSizePolicy,
    QFrame
)

import serial
from serial.tools import list_ports
HAS_PYSERIAL = True

# --------------------------- TDKPowerSupply ---------------------------
class TDKPowerSupply:
    """TDK电源控制类，通过RS232串口与电源通信"""

    # 类级别的串口锁，确保多个电源实例不会同时访问串口
    _serial_lock = threading.Lock()

    def __init__(self, address: int, serial_connection) -> None:
        self.address = address
        self.serial = serial_connection
        self.logger = logging.getLogger(f'TDKPowerSupply_{address}')
        # 设置日志级别为DEBUG以显示详细信息
        self.logger.setLevel(logging.DEBUG)
        # 每个电源实例有自己的操作锁
        self._instance_lock = threading.Lock()

    def connect(self) -> bool:
        """检查串口是否连接"""
        if self.serial and self.serial.is_open:
            return True
        return False

    def disconnect(self) -> None:
        """断开与电源的连接（在共享模式下，这只是一个逻辑操作）"""
        # 在共享串口模式下，我们不关闭串口
        pass

    def send_command(self, command: str, get_response: bool = True) -> Optional[str]:
        """
        发送SCPI命令到电源

        Args:
            command: SCPI命令字符串
            get_response: 是否需要读取响应

        Returns:
            如果需要响应，返回电源的响应字符串，否则返回None
        """
        if not self.serial or not self.serial.is_open:
            self.logger.error("串口未连接")
            return None

        # 使用类级别锁确保串口访问的线程安全
        with TDKPowerSupply._serial_lock:
            try:
                # 清空输入输出缓冲区
                self.serial.flushInput()
                self.serial.flushOutput()
                
                # 先选择设备地址
                select_cmd = f"INSTrument:NSELect {self.address}\n"
                self.serial.write(select_cmd.encode('ascii'))
                time.sleep(0.1)  # 增加设备选择等待时间
                
                # 发送实际命令
                full_command = f"{command}\n"
                self.serial.write(full_command.encode('ascii'))
                self.logger.debug(f"发送命令: {full_command.strip()}")
                
                # 如果需要响应，读取并返回
                if get_response:
                    time.sleep(0.15)  # 增加响应等待时间
                    
                    # 检查是否有数据可读
                    if self.serial.in_waiting > 0:
                        response = self.serial.read_until(b'\n').decode('ascii', errors='ignore').strip()
                        self.logger.debug(f"收到响应: '{response}'")
                        return response
                    else:
                        # 如果没有立即响应，再等待一下
                        time.sleep(0.1)
                        if self.serial.in_waiting > 0:
                            response = self.serial.read(self.serial.in_waiting).decode('ascii', errors='ignore').strip()
                            self.logger.debug(f"延迟收到响应: '{response}'")
                            return response
                        else:
                            self.logger.debug("无响应数据")
                            return ""
                else:
                    # 设置命令，等待一下让设备处理
                    time.sleep(0.05)
                
                return None
            except Exception as e:
                self.logger.error(f"地址{self.address}命令发送失败: {str(e)}")
                return None
    def set_voltage(self, voltage: float) -> bool:
        """
        设置输出电压

        Args:
            voltage: 电压值，单位V

        Returns:
            设置成功返回True，失败返回False
        """
        with self._instance_lock:
            self.logger.info(f"地址{self.address}: 设置电压 {voltage:.3f}V")
            try:
                cmd = f"VOLT:AMPL {voltage:.3f}"
                self.send_command(cmd, get_response=False)
                self.logger.debug(f"地址{self.address}: 发送电压命令 '{cmd}'")
                return True
            except Exception as e:
                self.logger.error(f"地址{self.address}: 电压设置异常: {e}")
                return False

    def set_current(self, current: float) -> bool:
        """
        设置输出电流

        Args:
            current: 电流值，单位A

        Returns:
            设置成功返回True，失败返回False
        """
        with self._instance_lock:
            self.logger.info(f"地址{self.address}: 设置电流 {current:.3f}A")
            try:
                cmd = f"CURR:AMPL {current:.3f}"
                self.send_command(cmd, get_response=False)
                self.logger.debug(f"地址{self.address}: 发送电流命令 '{cmd}'")
                return True
            except Exception as e:
                self.logger.error(f"地址{self.address}: 电流设置异常: {e}")
                return False

    def set_output(self, state: bool) -> bool:
        """
        控制电源输出开关

        Args:
            state: True为开启输出，False为关闭输出

        Returns:
            设置成功返回True，失败返回False
        """
        with self._instance_lock:
            status = 1 if state else 0
            self.logger.info(f"地址{self.address}: 设置输出{'开启' if state else '关闭'}")
            try:
                cmd = f"OUTP:STAT {status}"
                self.send_command(cmd, get_response=False)
                self.logger.debug(f"地址{self.address}: 发送输出命令 '{cmd}'")
                return True
            except Exception as e:
                self.logger.error(f"地址{self.address}: 输出设置异常: {e}")
                return False

    def get_voltage(self) -> Optional[float]:
        """获取当前设置的电压值"""
        response = self.send_command(":VOLT?", get_response=True)
        if response and response.strip():
            try:
                return float(response)
            except ValueError:
                self.logger.error(f"解析设置电压值失败: {response}")
        return None

    def get_current(self) -> Optional[float]:
        """获取当前设置的电流值"""
        response = self.send_command(":CURR?", get_response=True)
        if response and response.strip():
            try:
                return float(response)
            except ValueError:
                self.logger.error(f"解析设置电流值失败: {response}")
        return None

    def get_actual_voltage(self) -> Optional[float]:
        """获取实际输出电压值"""
        response = self.send_command("MEAS:VOLT?", get_response=True)
        if response:
            try:
                return float(response)
            except ValueError:
                self.logger.error(f"解析实际电压值失败: {response}")
        return None

    def get_actual_current(self) -> Optional[float]:
        """获取实际输出电流值"""
        response = self.send_command("MEAS:CURR?", get_response=True)
        if response:
            try:
                return float(response)
            except ValueError:
                self.logger.error(f"解析实际电流值失败: {response}")
        return None

    def get_output_status(self) -> Optional[bool]:
        """获取输出状态"""
        response = self.send_command("OUTP?", get_response=True)
        if response:
            try:
                return bool(int(response))
            except ValueError:
                self.logger.error(f"解析输出状态失败: {response}")
        return None

    def get_id(self) -> Optional[str]:
        """获取电源标识信息"""
        return self.send_command("*IDN?", get_response=True)


    def test_communication(self) -> bool:
        """测试与电源的基本通信"""
        try:
            # 尝试多个命令来测试通信
            test_commands = ["*IDN?", ":VOLT?", ":CURR?", "OUTP?"]
            
            for cmd in test_commands:
                self.logger.debug(f"测试命令: {cmd}")
                response = self.send_command(cmd, get_response=True)
                
                # 如果收到任何非空响应，说明通信成功
                if response and response.strip():
                    self.logger.info(f"地址{self.address}: 通信测试成功，命令 {cmd} 响应: {response}")
                    return True
            
            # 所有命令都没有响应
            self.logger.warning(f"地址{self.address}: 通信测试失败 - 所有命令无响应")
            return False
            
        except Exception as e:
            self.logger.error(f"地址{self.address}: 通信测试异常: {e}")
            return False

    def debug_power_settings(self) -> dict:
        """调试电源设置，返回所有状态信息"""
        debug_info = {}
        try:
            debug_info['device_id'] = self.get_id()
            debug_info['voltage_setting'] = self.get_voltage()
            debug_info['current_setting'] = self.get_current()
            debug_info['actual_voltage'] = self.get_actual_voltage()
            debug_info['actual_current'] = self.get_actual_current()
            debug_info['output_status'] = self.get_output_status()
            
            self.logger.info(f"地址{self.address} 调试信息: {debug_info}")
            return debug_info
        except Exception as e:
            self.logger.error(f"地址{self.address} 调试信息获取失败: {e}")
            return {}

    def reset(self) -> None:
        """重置电源到默认状态"""
        self.send_command("*RST", get_response=False)

# --------------------------- PowerThread ---------------------------
class PowerThread(QThread):
    # emits: voltage, current, power, timestamp, name
    data_signal = pyqtSignal(float, float, float, float, str)
    log_signal = pyqtSignal(str)
    status_signal = pyqtSignal(str)

    def __init__(self, name: str, poll_interval_ms: int = 500, widget=None):
        super().__init__()
        self.name = name
        self.poll_interval_ms = poll_interval_ms
        self._running = False
        self._target_v = 0.0
        self._target_i = 0.0
        self._power_supply = None
        self._widget = widget  # 存储对应的 PowerWidget 引用
        self._reading_paused = False  # 添加暂停读取标志

    def run(self):
        self._running = True
        self.log(f"{self.name}: 线程启动")

        # 使用已经建立的连接
        if not self._widget or not self._widget._power_supply:
            self.log(f"{self.name}: 电源未初始化，请先连接串口")
            return

        self._power_supply = self._widget._power_supply
        
        if self._power_supply.connect():
            try:
                idn = self._power_supply.get_id()
                if idn:
                    self.log(f"{self.name}: 已连接 (地址{self._power_supply.address}) {idn}")
            except:
                self.log(f"{self.name}: 已连接 (地址{self._power_supply.address})")
        else:
            self.log(f"{self.name}: 连接失败 (地址{self._power_supply.address})")
            return

        try:
            # 仅在开启输出时重置电源
            self.log(f"{self.name}: 正在重置电源 (地址{self._power_supply.address})...")
            self._power_supply.reset()
            time.sleep(0.2) # 等待重置完成

            self._power_supply.set_voltage(self._target_v)
            self._power_supply.set_current(self._target_i)
            self._power_supply.set_output(True)
            self.log(f"{self.name}: 设备初始化完成 (地址{self._power_supply.address})")
        except Exception as e:
            self.log(f"{self.name}: 初始设置失败: {e}")

        while self._running:
            ts = time.time()
            try:
                if self._power_supply and not self._reading_paused:
                    try:
                        voltage = self._power_supply.get_actual_voltage() or 0.0
                        current = self._power_supply.get_actual_current() or 0.0
                    except Exception as e:
                        self.log(f"{self.name}: 读取数据失败: {e}")
                        voltage = 0.0
                        current = 0.0
                else:
                    voltage = 0.0
                    current = 0.0

                # 根据不同电源计算功率
                if self.name == '射频电源':
                    power = voltage * 100  # 射频功率
                else:
                    power = voltage * current  # 偏压功率
                self.data_signal.emit(voltage, current, power, ts, self.name)
            except Exception as e:
                self.log(f"{self.name} ERROR: {e}")
                self.status_signal.emit(f"ERROR: {e}")

            self.msleep(self.poll_interval_ms)

        if self._power_supply:
            try:
                self._power_supply.set_output(False)
                self.log(f"{self.name}: 地址{self._power_supply.address} 已断开")
                self._power_supply.disconnect()
            except Exception as e:
                self.log(f"{self.name}: 断开时出错: {e}")
            finally:
                self._power_supply = None
        self.log(f"{self.name}: 线程停止")

    def stop(self):
        self._running = False
        self.wait(2000)

    def set_targets(self, voltage: float, current: float):
        self._target_v = voltage
        self._target_i = current
        self.log(f"{self.name}: 设置请求 - 电压{voltage:.3f}V 电流{current:.3f}A")
        
        if self._power_supply:
            try:
                self._reading_paused = True
                self.log(f"{self.name}: 暂停数据读取，开始设置参数...")
                
                self.log(f"{self.name}: 正在设置电压 {voltage:.3f}V...")
                voltage_success = self._power_supply.set_voltage(voltage)
                
                self.log(f"{self.name}: 正在设置电流 {current:.3f}A...")
                current_success = self._power_supply.set_current(current)
                
                self.log(f"{self.name}: 设置完成，恢复数据读取")
                self._reading_paused = False
                
                if voltage_success and current_success:
                    self.log(f"{self.name}: 设置完成 - 电压{voltage:.3f}V 电流{current:.3f}A")
                else:
                    self.log(f"{self.name}: 设置可能失败")
                    
            except Exception as e:
                self._reading_paused = False
                self.log(f"{self.name} 设置异常: {e}")

    # ------------------ helpers ------------------
    def log(self, msg: str):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_signal.emit(f"[{timestamp}] {msg}")

# --------------------------- PowerWidget (single side) ---------------------------
# --------------------------- PowerWidget (single side) ---------------------------
class PowerWidget(QWidget):
    set_targets_signal = pyqtSignal(float, float)

    def __init__(self, name: str, parent=None):
        super().__init__(parent)
        self.name = name
        self._power_supply = None
        self._build_ui()

    def _build_ui(self):
        title = QLabel(f"{self.name}")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet('font-weight: bold; font-size: 14px')

        # 串口通讯设置区（只保留地址）
        comm_group = QGroupBox('设备地址')
        comm_grid = QGridLayout()
        comm_grid.addWidget(QLabel('地址:'), 0, 0)
        self.edit_addr = QLineEdit('6' if self.name == '射频电源' else '7')
        # 设置地址为只读，固定不变
        self.edit_addr.setReadOnly(True)
        self.edit_addr.setStyleSheet('background-color: #f0f0f0; color: #666;')
        # 添加键盘事件处理
        self.edit_addr.keyPressEvent = lambda event: self._handle_lineedit_key_event(event, self.edit_addr)
        comm_grid.addWidget(self.edit_addr, 0, 1)
        comm_group.setLayout(comm_grid)

        # 输入设定区
        input_group = QGroupBox('输入设置')
        grid = QGridLayout()
        grid.addWidget(QLabel('电压 (V):'), 0, 0)
        self.spin_v = QDoubleSpinBox()
        self.spin_v.setRange(0, 1000)
        self.spin_v.setDecimals(3)
        self.spin_v.setSingleStep(0.1)
        # 添加键盘事件处理
        self.spin_v.keyPressEvent = lambda event: self._handle_spinbox_key_event(event, self.spin_v)
        grid.addWidget(self.spin_v, 0, 1)
        grid.addWidget(QLabel('电流 (A):'), 1, 0)
        self.spin_i = QDoubleSpinBox()
        self.spin_i.setRange(0, 100)
        self.spin_i.setDecimals(3)
        self.spin_i.setSingleStep(0.01)
        # 添加键盘事件处理
        self.spin_i.keyPressEvent = lambda event: self._handle_spinbox_key_event(event, self.spin_i)
        grid.addWidget(self.spin_i, 1, 1)
        
        # 添加电流限制提示
        current_limit_label = QLabel('最大电流0.6A，超出会设为0.6A')
        current_limit_label.setStyleSheet('font-size: 9px; color: #666;')
        grid.addWidget(current_limit_label, 2, 0, 1, 2)

        self.btn_set = QPushButton('设置')
        grid.addWidget(self.btn_set, 3, 0, 1, 2)
        input_group.setLayout(grid)

        # 输出显示区
        output_group = QGroupBox('电源返回值')
        form = QFormLayout()

        # 设备信息显示
        self.lbl_device_info = QLabel('未连接')
        self.lbl_device_info.setStyleSheet('color: #666; font-size: 10px;')
        self.lbl_device_info.setWordWrap(True)
        form.addRow(self.lbl_device_info)

        # 添加分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        form.addRow(line)

        self.lbl_v = QLabel('0.000 V')
        self.lbl_i = QLabel('0.000 A')
        self.lbl_p = QLabel('0.000 W')

        # 大号数字显示效果
        large_font_style = '''
            QLabel {
                font-size: 22px;
                font-weight: bold;
                color: #333;
            }
        '''
        self.lbl_v.setStyleSheet(large_font_style)
        self.lbl_i.setStyleSheet(large_font_style)
        self.lbl_p.setStyleSheet(large_font_style)

        # 固定宽度以保持对齐
        for lbl in [self.lbl_v, self.lbl_i, self.lbl_p]:
            lbl.setMinimumWidth(150)
            lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        form.addRow('电压:', self.lbl_v)
        form.addRow('电流:', self.lbl_i)
        if self.name == '射频电源':
            form.addRow('射频功率:', self.lbl_p)
        else:
            form.addRow('偏压功率:', self.lbl_p)
        output_group.setLayout(form)

        # 启动/停止按钮
        self.btn_start = QPushButton('启动')
        self.btn_stop = QPushButton('停止')
        self.btn_stop.setEnabled(False)
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_stop)

        vbox = QVBoxLayout()
        vbox.addWidget(title)
        vbox.addWidget(comm_group)
        vbox.addWidget(output_group)
        vbox.addWidget(input_group)
        vbox.addLayout(btn_layout)
        vbox.addStretch(1)
        self.setLayout(vbox)

        # 信号连接
        self.btn_set.clicked.connect(self._on_set)
        self.btn_start.clicked.connect(lambda: self.start_request(True))
        self.btn_stop.clicked.connect(lambda: self.start_request(False))

    def _handle_spinbox_key_event(self, event, spinbox):
        """处理SpinBox的键盘事件"""
        if event.key() in (Qt.Key_Return, Qt.Key_Enter, Qt.Key_Tab):
            # 触发设置按钮点击
            self._on_set()
            # 如果是Tab键，移动焦点到下一个控件
            if event.key() == Qt.Key_Tab:
                if spinbox == self.spin_v:
                    self.spin_i.setFocus()
                else:
                    self.btn_set.setFocus()
        else:
            # 调用原始的keyPressEvent
            QDoubleSpinBox.keyPressEvent(spinbox, event)

    def _handle_lineedit_key_event(self, event, lineedit):
        """处理LineEdit的键盘事件"""
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            # 地址输入框回车时重新连接设备
            self.parent().parent().update_device_connection()
        else:
            # 调用原始的keyPressEvent
            QLineEdit.keyPressEvent(lineedit, event)

    def _on_set(self):
        if not self._power_supply:
            QMessageBox.warning(self, "未连接", "请先连接电源设备")
            return
            
        v = float(self.spin_v.value())
        i = float(self.spin_i.value())
        
        # 限制最大电流为0.6A
        if i > 0.6:
            i = 0.6
            # 更新UI显示，让用户知道值被限制了
            self.spin_i.setValue(i)
            
        self.set_targets_signal.emit(v, i)

    def start_request(self, start: bool):
        if start and not self._power_supply:
            QMessageBox.warning(self, "未连接", "请先连接电源设备")
            return
            
        self.btn_start.setEnabled(not start)
        self.btn_stop.setEnabled(start)
        self.parent().parent().power_start_stop(self.name, start)

    def update_connection_status(self, connected: bool, serial_conn=None):
        """由MainWindow调用，更新连接状态"""
        if connected and serial_conn:
            try:
                addr = int(self.edit_addr.text().strip())
                self._power_supply = TDKPowerSupply(address=addr, serial_connection=serial_conn)
                
                # 测试基本通信
                if self._power_supply.test_communication():
                    # 尝试获取设备ID
                    idn = self._power_supply.get_id()
                    if idn and idn.strip():
                        self.lbl_device_info.setText(f'设备标识: \n{idn}')
                        self.lbl_device_info.setStyleSheet('color: #000;')
                        self.parent().parent().append_log(f"{self.name}: 连接成功 - {idn}")
                    else:
                        self.lbl_device_info.setText('已连接(标识获取失败)')
                        self.lbl_device_info.setStyleSheet('color: orange;')
                        self.parent().parent().append_log(f"{self.name}: 连接成功但无法获取设备标识")
                else:
                    self.lbl_device_info.setText('通信失败')
                    self.lbl_device_info.setStyleSheet('color: red;')
                    self.parent().parent().append_log(f"{self.name}: 地址{addr} 通信失败")
                    self._power_supply = None
                    
            except ValueError:
                self.lbl_device_info.setText('地址无效')
                self.lbl_device_info.setStyleSheet('color: red;')
                self._power_supply = None
        else:
            if self._power_supply:
                self._power_supply.disconnect()
            self._power_supply = None
            self.lbl_device_info.setText('未连接')
            self.lbl_device_info.setStyleSheet('color: #666;')

    # 通讯设置已集成到主界面，无需弹窗
    def _on_config(self):
        pass

    def update_outputs(self, v, i, p):
        self.lbl_v.setText(f"{v:.3f} V")
        self.lbl_i.setText(f"{i:.3f} A")
        self.lbl_p.setText(f"{p:.3f} W")

# --------------------------- CommConfigDialog ---------------------------
class CommConfigDialog(QDialog):
    def __init__(self, current=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle('通讯设置')
        self.resize(380, 180)
        self.current = current or {'mode': 'SIM', 'params': {}}
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout()
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(['SIM', 'SERIAL', 'TCP'])
        self.mode_combo.setCurrentText(self.current.get('mode', 'SIM'))
        layout.addWidget(QLabel('模式:'))
        layout.addWidget(self.mode_combo)

        # stack of simple param fields
        self.serial_port = QLineEdit(self.current.get('params', {}).get('port', ''))
        self.serial_port.keyPressEvent = lambda event: self._handle_lineedit_key_event(event, self.serial_port)
        self.serial_baud = QLineEdit(str(self.current.get('params', {}).get('baud', 9600)))
        self.serial_baud.keyPressEvent = lambda event: self._handle_lineedit_key_event(event, self.serial_baud)
        self.tcp_host = QLineEdit(self.current.get('params', {}).get('host', ''))
        self.tcp_host.keyPressEvent = lambda event: self._handle_lineedit_key_event(event, self.tcp_host)
        self.tcp_port = QLineEdit(str(self.current.get('params', {}).get('port', 5025)))
        self.tcp_port.keyPressEvent = lambda event: self._handle_lineedit_key_event(event, self.tcp_port)

        grid = QGridLayout()
        grid.addWidget(QLabel('串口名 (COM/ /dev):'), 0, 0)
        grid.addWidget(self.serial_port, 0, 1)
        grid.addWidget(QLabel('波特率:'), 1, 0)
        grid.addWidget(self.serial_baud, 1, 1)
        grid.addWidget(QLabel('TCP 主机:'), 2, 0)
        grid.addWidget(self.tcp_host, 2, 1)
        grid.addWidget(QLabel('TCP 端口:'), 3, 0)
        grid.addWidget(self.tcp_port, 3, 1)

        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def _handle_lineedit_key_event(self, event, lineedit):
        """处理LineEdit的键盘事件"""
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            # 回车时接受对话框
            self.accept()
        else:
            # 调用原始的keyPressEvent
            QLineEdit.keyPressEvent(lineedit, event)

    def get_settings(self):
        mode = self.mode_combo.currentText()
        params = {}
        if mode == 'SERIAL':
            params['port'] = self.serial_port.text().strip()
            try:
                params['baud'] = int(self.serial_baud.text().strip())
            except Exception:
                params['baud'] = 9600
        elif mode == 'TCP':
            params['host'] = self.tcp_host.text().strip()
            try:
                params['port'] = int(self.tcp_port.text().strip())
            except Exception:
                params['port'] = 5025
        return {'mode': mode, 'params': params}

# --------------------------- MainWindow ---------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('SIMS射频控制电源')
        self.resize(500, 700)
        self.shared_serial = None

        # 配置日志系统以显示调试信息
        logging.basicConfig(level=logging.DEBUG, 
                          format='%(levelname)s:%(name)s:%(message)s')

        central = QWidget()
        self.setCentralWidget(central)

        # left and right power widgets
        self.left_widget = PowerWidget('射频电源', parent=self)
        self.right_widget = PowerWidget('等离子体电源', parent=self)
        self.left_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        self.right_widget.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        # connect set signals
        self.left_widget.set_targets_signal.connect(lambda v, i: self._on_set('电源1', v, i))
        self.right_widget.set_targets_signal.connect(lambda v, i: self._on_set('电源2', v, i))

        # log area
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)

        # 串口连接区
        comm_group = QGroupBox('串口设置')
        comm_layout = QHBoxLayout()
        self.port_combo = QComboBox()
        self.update_ports()
        # 添加键盘事件处理
        self.port_combo.keyPressEvent = lambda event: self._handle_combo_key_event(event, self.port_combo)
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(['9600', '19200', '38400', '57600', '115200'])
        self.baud_combo.setCurrentText('115200')
        # 添加键盘事件处理
        self.baud_combo.keyPressEvent = lambda event: self._handle_combo_key_event(event, self.baud_combo)
        self.connect_btn = QPushButton('连接')
        self.refresh_btn = QPushButton('刷新')
        
        comm_layout.addWidget(QLabel("串口:"))
        comm_layout.addWidget(self.port_combo)
        comm_layout.addWidget(QLabel("波特率:"))
        comm_layout.addWidget(self.baud_combo)
        comm_layout.addWidget(self.refresh_btn)
        comm_layout.addWidget(self.connect_btn)
        comm_group.setLayout(comm_layout)

        self.connect_btn.clicked.connect(self.toggle_connection)
        self.refresh_btn.clicked.connect(self.update_ports)

        # layout
        main_layout = QGridLayout()
        main_layout.addWidget(comm_group, 0, 0, 1, 2) # 串口设置在最上面
        main_layout.addWidget(self.left_widget, 1, 0)
        main_layout.addWidget(self.right_widget, 1, 1)
        main_layout.addWidget(self.log_edit, 2, 0, 1, 2)  # 日志区域跨越两列
        main_layout.setRowStretch(1, 2)  # 让电源控制区域占更多空间
        main_layout.setRowStretch(2, 1)  # 日志区域占较少空间

        central.setLayout(main_layout)

        # threads dict
        self.threads = {}

        # menu
        self._build_menu()

        # timers for UI health (optional)
        self.ui_timer = QTimer(self)
        self.ui_timer.setInterval(1000)
        self.ui_timer.timeout.connect(lambda: None)
        self.ui_timer.start()

    def _handle_combo_key_event(self, event, combo):
        """处理ComboBox的键盘事件"""
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            # 下拉框回车时自动连接
            if not self.shared_serial or not self.shared_serial.is_open:
                self.toggle_connection()
        else:
            # 调用原始的keyPressEvent
            QComboBox.keyPressEvent(combo, event)

    def update_device_connection(self):
        """更新设备连接状态"""
        if self.shared_serial and self.shared_serial.is_open:
            self.left_widget.update_connection_status(True, self.shared_serial)
            self.right_widget.update_connection_status(True, self.shared_serial)

    def update_ports(self):
        self.port_combo.clear()
        if HAS_PYSERIAL:
            ports = [p.device for p in list_ports.comports()]
            self.port_combo.addItems(ports)

    def toggle_connection(self):
        if self.shared_serial and self.shared_serial.is_open:
            # 断开连接
            # 先停止所有线程
            for name in list(self.threads.keys()):
                self.power_start_stop(name, False)

            self.shared_serial.close()
            self.shared_serial = None
            self.connect_btn.setText('连接')
            self.port_combo.setEnabled(True)
            self.baud_combo.setEnabled(True)
            self.refresh_btn.setEnabled(True)
            self.left_widget.update_connection_status(False)
            self.right_widget.update_connection_status(False)
            self.append_log("串口已断开")
        else:
            # 连接
            port = self.port_combo.currentText()
            baud = int(self.baud_combo.currentText())
            if not port:
                QMessageBox.warning(self, "警告", "没有可用的串口")
                return
            try:
                self.shared_serial = serial.Serial(port, baud, timeout=1)
                # 移除串口连接延迟，立即继续
                self.connect_btn.setText('断开')
                self.port_combo.setEnabled(False)
                self.baud_combo.setEnabled(False)
                self.refresh_btn.setEnabled(False)
                
                # 立即更新设备连接状态，不添加延迟
                self.left_widget.update_connection_status(True, self.shared_serial)
                self.right_widget.update_connection_status(True, self.shared_serial)
                self.append_log(f"串口 {port} 已连接")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法打开串口 {port}: {e}")
                self.shared_serial = None

    def _build_menu(self):
        menubar = self.menuBar()
        filem = menubar.addMenu('文件')
        save_log_act = QAction('保存日志', self)
        save_log_act.triggered.connect(self.save_log)
        filem.addAction(save_log_act)

        exit_act = QAction('退出', self)
        exit_act.triggered.connect(self.close)
        filem.addAction(exit_act)

        helpm = menubar.addMenu('帮助')
        about_act = QAction('关于', self)
        about_act.triggered.connect(self._about)
        helpm.addAction(about_act)

    def _about(self):
        QMessageBox.information(self, '关于', '电源控制界面')

    def append_log(self, text: str):
        self.log_edit.append(text)

    # called from PowerWidget when start/stop pressed
    def power_start_stop(self, name: str, start: bool):
        if start:
            widget = self.left_widget if name == '电源1' else self.right_widget
            
            if not widget._power_supply:
                error_msg = "请先连接电源设备"
                QMessageBox.warning(self, "未连接", error_msg)
                self.append_log(f"{name}: 启动失败 - {error_msg}")
                widget.btn_start.setEnabled(True)
                widget.btn_stop.setEnabled(False)
                return
                
            # 启动线程
            th = PowerThread(name=name, widget=widget)
            self.append_log(f"{name}: 启动")
            
            # 连接信号并启动线程 - 移除所有同步操作，立即启动
            th.data_signal.connect(self._on_data)
            th.log_signal.connect(self.append_log)
            th.status_signal.connect(lambda s: self.append_log(f"[{name}] {s}"))
            
            # 立即启动线程，不等待
            th.start()
            self.threads[name] = th
            
            # 立即返回，不做任何等待
            return
        else:
            th = self.threads.get(name)
            if th:
                th.stop()
                del self.threads[name]

    def _on_set(self, name: str, v: float, i: float):
        widget = self.left_widget if name == '电源1' else self.right_widget
        if not widget._power_supply:
            QMessageBox.warning(self, "未连接", "请先连接电源设备")
            return
            
        self.append_log(f"{name}: 设置 电压{v:.3f}V 电流{i:.3f}A")
        
        th = self.threads.get(name)
        if th:
            th.set_targets(v, i)
        else:
            self.append_log(f"{name}: 未启动，请先启动电源")
            QMessageBox.warning(self, "未启动", "请先启动电源再进行设置")

    def _on_data(self, v, i, p, ts, name):
        # 更新UI显示，不输出数据日志（避免日志刷屏）
        if name == '射频电源':
            self.left_widget.update_outputs(v, i, p)
        elif name == '等离子体电源':
            self.right_widget.update_outputs(v, i, p)

    def save_log(self):
        path, _ = QFileDialog.getSaveFileName(self, '保存日志', os.getcwd(), 'Text Files (*.txt);;All Files (*)')
        if path:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self.log_edit.toPlainText())
            QMessageBox.information(self, '保存', f'日志已保存到 {path}')

    def closeEvent(self, event):
        # stop threads
        for name, th in list(self.threads.items()):
            try:
                th.stop()
            except Exception:
                pass
        if self.shared_serial and self.shared_serial.is_open:
            self.shared_serial.close()
        event.accept()

# --------------------------- main ---------------------------
def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

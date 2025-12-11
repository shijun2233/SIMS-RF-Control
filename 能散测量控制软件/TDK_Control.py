
import time
import logging
import threading
from typing import Optional, Union
from collections import deque
from datetime import datetime
import logging
import threading

import serial
from serial.tools import list_ports



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
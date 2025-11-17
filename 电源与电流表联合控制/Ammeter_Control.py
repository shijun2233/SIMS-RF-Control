import serial
import time
from typing import Optional, Union, List
from serial import SerialException
from enum import Enum

# 设备常量定义（源自用户手册）
MAX_CURRENT_6485 = 0.021  # 6485最大电流21mA
MAX_CURRENT_6487 = 0.021  # 6487最大电流21mA
MAX_VOLTAGE_6485 = 220    # 6485最大输入电压220V峰值
MAX_VOLTAGE_6487 = 505    # 6487最大输入电压505V峰值
VOLTAGE_RANGES_6487 = [10, 50, 500]  # 6487电压源量程
CURRENT_CLAMPS_6487 = [2.5e-5, 2.5e-4, 2.5e-3, 2.5e-2]  # 6487电流钳位选项
MEASURE_RANGES = [2e-9, 2e-8, 2e-7, 2e-6, 2e-5, 2e-4, 2e-3, 2.1e-2]  # 8个电流量程（2nA~21mA）

class InstrumentModel(Enum):
    """仪器型号枚举"""
    MODEL_6485 = "6485"
    MODEL_6487 = "6487"

class FilterType(Enum):
    """滤波器类型枚举"""
    MEDIAN = "MED"
    AVERAGE = "AVER"

class KeithleyPicoammeter:
    """Keithley 6485/6487 皮安表SCPI控制库"""
    
    def __init__(
        self,
        port: str,
        baudrate: int = 9600,
        model: InstrumentModel = InstrumentModel.MODEL_6485,
        timeout: float = 1.0
    ):
        """
        初始化设备连接
        :param port: 串口名称（如COM3、/dev/ttyUSB0）
        :param baudrate: 波特率（默认9600，需与设备一致）
        :param model: 仪器型号
        :param timeout: 串口超时时间（秒）
        """
        self.port = port
        self.baudrate = baudrate
        self.model = model
        self.timeout = timeout
        self.serial: Optional[serial.Serial] = None
        self._is_connected = False
        self._current_range: Optional[float] = None  # 当前量程
        self._filter_enabled = False  # 滤波器启用状态

    def connect(self) -> bool:
        """建立串口连接"""
        try:
            self.serial = serial.Serial(
                port=self.port,
                baudrate=self.baudrate,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                bytesize=serial.EIGHTBITS,
                timeout=self.timeout,
                xonxoff=False,
                rtscts=False
            )
            time.sleep(0.5)  # 等待设备初始化
            self._is_connected = True
            # 复位设备到默认状态
            self.send_command("*RST")
            time.sleep(0.2)
            # 确认设备标识
            idn = self.get_identity()
            if self.model.value in idn:
                print(f"成功连接 {idn}")
                return True
            else:
                print(f"设备型号不匹配，实际设备：{idn}")
                self.disconnect()
                return False
        except SerialException as e:
            print(f"连接失败：{str(e)}")
            return False

    def disconnect(self) -> None:
        """断开串口连接"""
        if self.serial and self.serial.is_open:
            # 关闭电压源（仅6487）
            if self.model == InstrumentModel.MODEL_6487:
                self.set_voltage_source_status(False)
            self.serial.close()
        self._is_connected = False
        print("已断开连接")

    def send_command(self, cmd: str, get_response: bool = False) -> Optional[str]:
        """
        发送SCPI命令
        :param cmd: SCPI命令字符串
        :param get_response: 是否需要返回响应
        :return: 响应结果（None表示无响应或失败）
        """
        if not self._is_connected or not self.serial:
            print("未建立连接")
            return None
        
        try:
            # 清空缓冲区
            self.serial.flushInput()
            self.serial.flushOutput()
            # 发送命令（末尾添加换行符）
            cmd = cmd.strip() + "\n"
            self.serial.write(cmd.encode("ascii"))
            time.sleep(0.1)  # 等待命令执行
            
            if get_response:
                response = b""
                while True:
                    data = self.serial.read(1)
                    if not data:
                        break
                    response += data
                    if response.endswith(b"\n"):
                        break
                response_str = response.decode("ascii", errors="ignore").strip()
                # 检查错误响应
                if response_str.startswith("ERROR"):
                    print(f"命令执行错误 [{cmd.strip()}]: {response_str}")
                    return None
                return response_str
            return None
        except SerialException as e:
            print(f"命令发送失败 [{cmd.strip()}]: {str(e)}")
            return None

    def get_identity(self) -> Optional[str]:
        """获取设备标识（*IDN?）"""
        return self.send_command("*IDN?", get_response=True)

    def reset(self) -> bool:
        """复位设备到出厂默认状态"""
        result = self.send_command("*RST")
        time.sleep(0.5)
        return result is not None

    # ------------------------------ 电流测量相关 ------------------------------
    def set_current_range(self, range_val: float) -> bool:
        """
        设置电流量程
        :param range_val: 量程值（必须在MEASURE_RANGES中）
        """
        if range_val not in MEASURE_RANGES:
            print(f"无效量程，可选量程：{MEASURE_RANGES}")
            return False
        # 检查电压安全限制
        if (self.model == InstrumentModel.MODEL_6485 and range_val >= 2e-3) or \
           (self.model == InstrumentModel.MODEL_6487 and range_val >= 2e-3):
            print(f"警告：{range_val}A量程下输入电压不得超过60V")
        
        cmd = f"RANG {range_val:.9f}"
        result = self.send_command(cmd)
        if result is not None:
            self._current_range = range_val
            print(f"已设置量程：{range_val}A")
            return True
        return False

    def set_auto_range(self, enable: bool = True) -> bool:
        """启用/禁用自动量程"""
        cmd = f"RANG:AUTO {'ON' if enable else 'OFF'}"
        result = self.send_command(cmd)
        if result is not None:
            print(f"{'启用' if enable else '禁用'}自动量程")
            return True
        return False

    def measure_current(self) -> Optional[float]:
        """测量电流（返回实际测量值）"""
        # 确保电流功能被选中（仅6487需要）
        if self.model == InstrumentModel.MODEL_6487:
            self.send_command("FUNC 'CURR'")
            time.sleep(0.1)
        # 触发测量并返回结果
        response = self.send_command("READ?", get_response=True)
        if response is not None:
            try:
                return float(response)
            except ValueError:
                print(f"电流值解析失败：{response}")
        return None

    def set_zero_check(self, enable: bool = True) -> bool:
        """启用/禁用零点检查（连接/断开输入时使用）"""
        cmd = f"SYST:ZCH {'ON' if enable else 'OFF'}"
        result = self.send_command(cmd)
        if result is not None:
            print(f"{'启用' if enable else '禁用'}零点检查")
            return True
        return False

    def set_zero_correct(self, enable: bool = True) -> bool:
        """启用/禁用零点校正（抵消偏移）"""
        if enable:
            # 先激活零点检查，获取校正值
            self.set_zero_check(True)
            time.sleep(0.2)
            self.send_command("SYST:ZCOR:ACQ")
            time.sleep(0.1)
            cmd = "SYST:ZCOR ON"
        else:
            cmd = "SYST:ZCOR OFF"
        
        result = self.send_command(cmd)
        if result is not None:
            print(f"{'启用' if enable else '禁用'}零点校正")
            return True
        return False

    # ------------------------------ 6487专属：电压源控制 ------------------------------
    def set_voltage_source_range(self, range_val: int) -> bool:
        """
        设置电压源量程（仅6487）
        :param range_val: 量程值（10/50/500V）
        """
        if self.model != InstrumentModel.MODEL_6487:
            print("该功能仅支持6487型号")
            return False
        
        if range_val not in VOLTAGE_RANGES_6487:
            print(f"无效电压量程，可选：{VOLTAGE_RANGES_6487}")
            return False
        
        cmd = f"SOUR:VOLT:RANG {range_val}"
        result = self.send_command(cmd)
        if result is not None:
            print(f"已设置电压源量程：{range_val}V")
            return True
        return False

    def set_voltage(self, voltage: float) -> bool:
        """
        设置电压源输出电压（仅6487）
        :param voltage: 输出电压（-505~505V，需在选定量程内）
        """
        if self.model != InstrumentModel.MODEL_6487:
            print("该功能仅支持6487型号")
            return False
        
        # 检查电压范围
        current_range = self.get_voltage_source_range()
        if current_range:
            max_volt = current_range
            if abs(voltage) > max_volt:
                print(f"电压超出量程（最大{max_volt}V）")
                return False
        
        cmd = f"SOUR:VOLT {voltage:.3f}"
        result = self.send_command(cmd)
        if result is not None:
            print(f"已设置电压源输出：{voltage}V")
            return True
        return False

    def set_voltage_source_current_clamp(self, clamp_current: float) -> bool:
        """
        设置电压源电流钳位（仅6487）
        :param clamp_current: 钳位电流（25μA/250μA/2.5mA/25mA）
        """
        if self.model != InstrumentModel.MODEL_6487:
            print("该功能仅支持6487型号")
            return False
        
        if clamp_current not in CURRENT_CLAMPS_6487:
            print(f"无效钳位电流，可选：{CURRENT_CLAMPS_6487}")
            return False
        
        cmd = f"SOUR:VOLT:ILIM {clamp_current:.6f}"
        result = self.send_command(cmd)
        if result is not None:
            print(f"已设置电流钳位：{clamp_current}A")
            return True
        return False

    def set_voltage_source_status(self, enable: bool = True) -> bool:
        """
        开启/关闭电压源输出（仅6487）
        :param enable: True=开启，False=关闭
        """
        if self.model != InstrumentModel.MODEL_6487:
            print("该功能仅支持6487型号")
            return False
        
        cmd = f"SOUR:VOLT:STAT {'ON' if enable else 'OFF'}"
        result = self.send_command(cmd)
        if result is not None:
            print(f"电压源{'已开启' if enable else '已关闭'}")
            return True
        return False

    def get_voltage_source_range(self) -> Optional[float]:
        """获取当前电压源量程（仅6487）"""
        if self.model != InstrumentModel.MODEL_6487:
            print("该功能仅支持6487型号")
            return None
        return self.send_command("SOUR:VOLT:RANG?", get_response=True)

    # ------------------------------ 6487专属：欧姆测量 ------------------------------
    def measure_resistance(self, voltage: float = 10.0, clamp_current: float = 2.5e-3) -> Optional[float]:
        """
        测量电阻（仅6487）
        :param voltage: 测试电压（需在量程内）
        :param clamp_current: 电流钳位值
        :return: 电阻值（Ω）
        """
        if self.model != InstrumentModel.MODEL_6487:
            print("该功能仅支持6487型号")
            return None
        
        try:
            # 步骤1：设置电压源
            self.set_voltage_source_range(self._get_nearest_voltage_range(voltage))
            self.set_voltage(voltage)
            self.set_voltage_source_current_clamp(clamp_current)
            
            # 步骤2：零点校正
            self.set_zero_check(True)
            self.set_zero_correct(True)
            self.set_zero_check(False)
            
            # 步骤3：激活欧姆功能
            self.send_command("SENS:OHMS ON")
            time.sleep(0.2)
            
            # 步骤4：开启电压源并测量
            self.set_voltage_source_status(True)
            time.sleep(0.3)  # 稳定时间
            response = self.send_command("READ?", get_response=True)
            self.set_voltage_source_status(False)
            
            if response:
                resistance = float(response)
                print(f"测量电阻：{resistance:.2f}Ω")
                return resistance
            return None
        except Exception as e:
            print(f"电阻测量失败：{str(e)}")
            self.set_voltage_source_status(False)
            return None

    # ------------------------------ 滤波器控制 ------------------------------
    def set_filter(self, filter_type: FilterType, enable: bool = True, param: int = 5) -> bool:
        """
        设置滤波器（中值/平均）
        :param filter_type: 滤波器类型
        :param enable: 是否启用
        :param param: 参数（中值：1-5阶；平均：2-100次）
        """
        if not (1 <= param <= 5) and filter_type == FilterType.MEDIAN:
            print("中值滤波器参数范围：1-5")
            return False
        if not (2 <= param <= 100) and filter_type == FilterType.AVERAGE:
            print("平均滤波器参数范围：2-100")
            return False
        
        try:
            if filter_type == FilterType.MEDIAN:
                self.send_command(f"MED:RANK {param}")
                self.send_command(f"MED {'ON' if enable else 'OFF'}")
            else:
                self.send_command(f"AVER:COUN {param}")
                self.send_command(f"AVER:TCON MOV")  # 移动平均模式
                self.send_command(f"AVER {'ON' if enable else 'OFF'}")
            
            self._filter_enabled = enable
            print(f"{'启用' if enable else '禁用'} {filter_type.name}滤波器（参数：{param}）")
            return True
        except Exception as e:
            print(f"滤波器设置失败：{str(e)}")
            return False

    # ------------------------------ 缓冲区操作 ------------------------------
    def configure_buffer(self, sample_count: int = 1000) -> bool:
        """
        配置数据缓冲区
        :param sample_count: 采样点数（6485最大2500，6487最大3000）
        """
        max_samples = 2500 if self.model == InstrumentModel.MODEL_6485 else 3000
        if sample_count < 1 or sample_count > max_samples:
            print(f"无效采样点数，最大支持{max_samples}点")
            return False
        
        # 配置缓冲区
        self.send_command(f"TRAC:POIN {sample_count}")
        self.send_command("TRAC:CLE")  # 清空缓冲区
        self.send_command("TRAC:FEED:CONT NEXT")  # 从下一个读数开始存储
        print(f"缓冲区已配置：{sample_count}点")
        return True

    def capture_buffer_data(self, sample_count: int = 1000, range_val: float = 2e-6) -> Optional[List[float]]:
        """
        高速采集数据到缓冲区
        :param sample_count: 采样点数
        :param range_val: 测量量程
        :return: 采集的数据列表
        """
        if not self.configure_buffer(sample_count):
            return None
        
        try:
            # 配置高速测量参数
            self.set_current_range(range_val)
            self.send_command("NPLC .01")  # 快速积分率（0.01PLC）
            self.send_command("SYST:AZER:STAT OFF")  # 关闭自动清零
            self.send_command("DISP:ENAB OFF")  # 关闭显示器节省资源
            self.send_command(f"TRIG:COUN {sample_count}")  # 设置触发次数
            self.send_command("TRIG:DEL 0")  # 触发延迟0秒
            
            # 开始采集
            print("开始高速采集...")
            self.send_command("INIT")
            # 等待采集完成（根据采样点数估算时间）
            wait_time = sample_count * 0.001 + 1  # 1ms/点 + 缓冲时间
            time.sleep(wait_time)
            
            # 读取缓冲区数据
            response = self.send_command("TRAC:DATA?", get_response=True)
            # 恢复显示
            self.send_command("DISP:ENAB ON")
            
            if response:
                # 解析数据（逗号分隔）
                data = [float(x.strip()) for x in response.split(",") if x.strip()]
                print(f"采集完成，实际获取{len(data)}个数据点")
                return data[:sample_count]  # 确保不超过请求点数
            return None
        except Exception as e:
            print(f"缓冲区采集失败：{str(e)}")
            self.send_command("DISP:ENAB ON")
            return None

    # ------------------------------ 辅助函数 ------------------------------
    def _get_nearest_voltage_range(self, voltage: float) -> int:
        """获取最接近的电压量程（仅6487）"""
        voltage = abs(voltage)
        for range_val in sorted(VOLTAGE_RANGES_6487):
            if range_val >= voltage:
                return range_val
        return VOLTAGE_RANGES_6487[-1]  # 返回最大量程

    def get_error_status(self) -> Optional[str]:
        """获取设备错误状态"""
        return self.send_command("SYST:ERR?", get_response=True)

    def self_test(self) -> bool:
        """设备自检"""
        response = self.send_command("*TST?", get_response=True)
        if response == "0":
            print("自检通过")
            return True
        else:
            print(f"自检失败，错误码：{response}")
            return False



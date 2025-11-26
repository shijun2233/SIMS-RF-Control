# app_window.py
import sys
import json
import csv
import time
import threading
from datetime import datetime

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                             QVBoxLayout, QGridLayout, QGroupBox, QLabel,
                             QLineEdit, QPushButton, QMessageBox, QFileDialog,
                             QTextEdit)
from PyQt5.QtCore import Qt, pyqtSignal, QObject
from PyQt5.QtGui import QFont

import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# 假设原控制类在同一目录
from TDK_Control import TDKPowerSupply
from Ammeter_Control import KeithleyPicoammeter


# ------------ 线程通信 ------------
class SigEmitter(QObject):
    append_data = pyqtSignal(tuple)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('SIMS RF 控制测量界面')
        self.resize(640, 360)
        self._load_geometry()

        # 设备 & 数据
        self.tdk = None
        self.amm = None
        self.data = []
        self._stop_event = threading.Event()

        # 信号
        self.sig = SigEmitter()
        self.sig.append_data.connect(self._on_append_data)

        # 构建 UI
        self._build_ui()

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_lo = QVBoxLayout(central)

        # 顶部为控制区，随后是参数区，底部为图表与输出终端左右两栏
        main_lo.addLayout(self._top_bar())
        main_lo.addLayout(self._bottom_bar())
        h = QHBoxLayout()
        self.canvas = self._build_canvas()
        h.addWidget(self.canvas, 3)
        # 输出终端（只读）
        self.terminal = QTextEdit()
        self.terminal.setReadOnly(True)
        h.addWidget(self.terminal, 1)
        main_lo.addLayout(h)

    # ---------- 顶部工具栏 ----------
    def _top_bar(self):
        lo = QHBoxLayout()

        # TDK
        g1 = QGroupBox('TDK 电源')
        g1.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        grid = QGridLayout(g1)
        self.pwr_port = QLineEdit('COM11')
        self.pwr_addr = QLineEdit('6')
        self.voltage_entry = QLineEdit()
        self.current_entry = QLineEdit()
        grid.addWidget(QLabel('串口'), 0, 0)
        grid.addWidget(self.pwr_port, 0, 1)
        grid.addWidget(QLabel('地址'), 1, 0)
        grid.addWidget(self.pwr_addr, 1, 1)
        grid.addWidget(QLabel('电压(V)'), 0, 2)
        grid.addWidget(self.voltage_entry, 0, 3)
        grid.addWidget(QLabel('电流(A)'), 1, 2)
        grid.addWidget(self.current_entry, 1, 3)

        btn_connect_p = QPushButton('连接电源')
        btn_disconnect_p = QPushButton('断开电源')
        btn_set_v = QPushButton('设置电压')
        btn_set_i = QPushButton('设置电流')
        btn_on = QPushButton('输出 ON')
        btn_off = QPushButton('输出 OFF')
        btn_connect_p.clicked.connect(self.connect_power)
        btn_disconnect_p.clicked.connect(self.disconnect_power)
        btn_set_v.clicked.connect(self.set_voltage)
        btn_set_i.clicked.connect(self.set_current)
        btn_on.clicked.connect(lambda: self.set_output(True))
        btn_off.clicked.connect(lambda: self.set_output(False))

        grid.addWidget(btn_connect_p, 2, 0)
        grid.addWidget(btn_disconnect_p, 2, 1)
        grid.addWidget(btn_set_v, 0, 4)
        grid.addWidget(btn_set_i, 1, 4)
        grid.addWidget(btn_on, 2, 2)
        grid.addWidget(btn_off, 2, 3)
        lo.addWidget(g1)

        # Keithley
        g2 = QGroupBox('Keithley 皮安表')
        g2.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        grid2 = QGridLayout(g2)
        self.amm_port = QLineEdit('COM12')
        grid2.addWidget(QLabel('串口'), 0, 0)
        grid2.addWidget(self.amm_port, 0, 1)
        btn_connect_a = QPushButton('连接安培表')
        btn_disconnect_a = QPushButton('断开安培表')
        btn_src = QPushButton('选择电源测量')
        btn_prep = QPushButton('准备测量')
        btn_connect_a.clicked.connect(self.connect_amm)
        btn_disconnect_a.clicked.connect(self.disconnect_amm)
        btn_src.clicked.connect(self.select_source_measure)
        btn_prep.clicked.connect(self.prepare_measure)
        grid2.addWidget(btn_connect_a, 1, 0)
        grid2.addWidget(btn_disconnect_a, 1, 1)
        grid2.addWidget(btn_src, 2, 0)
        grid2.addWidget(btn_prep, 2, 1)
        lo.addWidget(g2)

        lo.addStretch()
        btn_save = QPushButton('保存数据')
        btn_save.clicked.connect(self.save_data)
        btn_stop = QPushButton('停止全部')
        btn_stop.clicked.connect(self.stop_operations)
        lo.addWidget(btn_save)
        lo.addWidget(btn_stop)
        return lo

    # ---------- 底部参数 ----------
    def _bottom_bar(self):
        lo = QHBoxLayout()

        # 步进
        g3 = QGroupBox('步进输出参数')
        g3.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        grid3 = QGridLayout(g3)
        self.start_v, self.stop_v, self.step_v, self.step_time = [QLineEdit() for _ in range(4)]
        self.step_time.setText('0.2')
        labels = ('起始 V', '终止 V', '步长 V', '每步时间(s)')
        for i, (lab, w) in enumerate(zip(labels, (self.start_v, self.stop_v, self.step_v, self.step_time))):
            grid3.addWidget(QLabel(lab), 0, i * 2)
            grid3.addWidget(w, 0, i * 2 + 1)
        btn_go = QPushButton('开始阶梯输出并测量')
        btn_go.clicked.connect(self.start_step_and_measure)
        grid3.addWidget(btn_go, 1, 0, 1, 8)
        lo.addWidget(g3)

        # 连续测量
        g4 = QGroupBox('连续测量参数')
        g4.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        grid4 = QGridLayout(g4)
        self.measure_steps = QLineEdit('10')
        self.measure_interval = QLineEdit('0.2')
        grid4.addWidget(QLabel('测量步数'), 0, 0)
        grid4.addWidget(self.measure_steps, 0, 1)
        grid4.addWidget(QLabel('间隔(s)'), 0, 2)
        grid4.addWidget(self.measure_interval, 0, 3)
        h = QHBoxLayout()
        btn_m = QPushButton('开始测量')
        btn_m.clicked.connect(self.start_measure)
        btn_s = QPushButton('单次测量')
        btn_s.clicked.connect(self.single_measure)
        h.addWidget(btn_m)
        h.addWidget(btn_s)
        grid4.addLayout(h, 1, 0, 1, 4)
        lo.addWidget(g4)
        return lo

    # ---------- 绘图 ----------
    def _build_canvas(self):
        self.fig = Figure(figsize=(8, 5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_xlabel('Voltage (V)')
        self.ax.set_ylabel('Current (A)')
        self.line, = self.ax.plot([], [], '-o', markersize=4)
        return FigureCanvas(self.fig)

    # ---------------- 原函数名完全一致 ----------------
    def connect_power(self):
        port = self.pwr_port.text().strip()
        addr = int(self.pwr_addr.text().strip() or '1')
        if not port:
            return QMessageBox.warning(self, '提示', '请填写电源串口')
        try:
            import serial
            ser = serial.Serial(port, baudrate=9600, timeout=0.5)
        except Exception as e:
            return QMessageBox.critical(self, '串口打开失败', str(e))
        self.tdk = TDKPowerSupply(addr, ser)
        ok = self.tdk.test_communication()
        QMessageBox.information(self, '连接电源', '通信成功' if ok else '通信可能失败，请检查')

    def disconnect_power(self):
        if self.tdk:
            self.tdk.disconnect()
            self.tdk = None
            QMessageBox.information(self, '断开', '电源已断开')

    def connect_amm(self):
        port = self.amm_port.text().strip()
        if not port:
            return QMessageBox.warning(self, '提示', '请填写安培表串口')
        self.amm = KeithleyPicoammeter(port)
        ok = self.amm.connect()
        QMessageBox.information(self, '连接安培表', '连接成功' if ok else '连接失败')

    def disconnect_amm(self):
        if self.amm:
            self.amm.disconnect()
            self.amm = None
            QMessageBox.information(self, '断开', '安培表已断开')

    def set_voltage(self):
        if not self.tdk:
            return QMessageBox.warning(self, '未连接', '请先连接电源')
        try:
            v = float(self.voltage_entry.text())
        except Exception:
            return QMessageBox.critical(self, '错误', '无效电压值')
        self.tdk.set_voltage(v)

    def set_current(self):
        if not self.tdk:
            return QMessageBox.warning(self, '未连接', '请先连接电源')
        try:
            i = float(self.current_entry.text())
        except Exception:
            return QMessageBox.critical(self, '错误', '无效电流值')
        self.tdk.set_current(i)

    def set_output(self, state: bool):
        if not self.tdk:
            return QMessageBox.warning(self, '未连接', '请先连接电源')
        self.tdk.set_output(state)

    def select_source_measure(self):
        # 执行安培表的 start_current_measurement 方法（如果已连接）
        if not self.amm:
            return QMessageBox.warning(self, '未连接', '请先连接安培表')
        ok = False
        try:
            ok = self.amm.start_current_measurement()
        except Exception as e:
            QMessageBox.critical(self, '错误', f'start_current_measurement 失败: {e}')
            try:
                self.log(f'start_current_measurement exception: {e}')
            except Exception:
                pass
            return
        if ok:
            QMessageBox.information(self, '选择电源测量', '已切换到电流测量（start_current_measurement 返回 True）')
            try:
                self.log('select_source_measure: start_current_measurement -> True')
            except Exception:
                pass
        else:
            QMessageBox.warning(self, '选择电源测量', 'start_current_measurement 返回 False')
            try:
                self.log('select_source_measure: start_current_measurement -> False')
            except Exception:
                pass

    def prepare_measure(self):
        if not self.amm:
            return QMessageBox.warning(self, '未连接', '请先连接安培表')
        cmds = ["*RST", "SYST:ACH ON", "RANG 2e-9", "INIT", "SYST:ZCOR:ACQ",
                "SYST:ZCOR ON", "RANG:AUTO ON", "SYST:ZCH OFF"]
        for c in cmds:
            self.amm.send_command(c)
            time.sleep(0.05)
        QMessageBox.information(self, '准备', '已发送准备测量命令')
        try:
            self.log('prepare_measure: sent preparation commands')
        except Exception:
            pass

    def single_measure(self):
        if not self.amm:
            return QMessageBox.warning(self, '未连接', '请先连接安培表')
        val = self.amm.measure_current()
        if val is None:
            return QMessageBox.critical(self, '测量失败', '未能读取电流')
        volt = self.tdk.get_actual_voltage() if self.tdk else None
        self.sig.append_data.emit((volt, val, datetime.now().isoformat()))

    def start_measure(self):
        if not self.amm:
            return QMessageBox.warning(self, '未连接', '请先连接安培表')
        try:
            steps = int(self.measure_steps.text())
            interval = float(self.measure_interval.text())
        except Exception:
            return QMessageBox.critical(self, '错误', '请填写有效的步数与间隔')
        self._stop_event.clear()
        threading.Thread(target=self._measure_loop, args=(steps, interval), daemon=True).start()

    def _measure_loop(self, steps, interval):
        for _ in range(steps):
            if self._stop_event.is_set():
                break
            val = self.amm.measure_current()
            volt = self.tdk.get_actual_voltage() if self.tdk else None
            self.sig.append_data.emit((volt, val, datetime.now().isoformat()))
            time.sleep(interval)

    def start_step_and_measure(self):
        if not (self.tdk and self.amm):
            return QMessageBox.warning(self, '未连接', '请先连接电源与安培表')
        try:
            start, stop, step, step_time = map(float, (self.start_v.text(), self.stop_v.text(),
                                                       self.step_v.text(), self.step_time.text()))
        except Exception:
            return QMessageBox.critical(self, '错误', '请填写有效的步进参数')
        if (stop - start) * step < 0:
            return QMessageBox.critical(self, '错误', '步长方向与起止不匹配')
        self._stop_event.clear()
        threading.Thread(target=self._step_and_measure_thread,
                         args=(start, stop, step, step_time), daemon=True).start()

    def _step_and_measure_thread(self, start, stop, step, step_time):
        volt, ascending = start, step > 0
        while True:
            if (ascending and volt > stop) or (not ascending and volt < stop):
                break
            if self._stop_event.is_set():
                break
            self.tdk.set_voltage(volt)
            time.sleep(0.2)
            cur = self.amm.measure_current()
            self.sig.append_data.emit((volt, cur, datetime.now().isoformat()))
            elapsed = 0.0
            while elapsed < step_time:
                if self._stop_event.is_set():
                    break
                time.sleep(0.05)
                elapsed += 0.05
            volt += step

    def stop_operations(self):
        self._stop_event.set()
        QMessageBox.information(self, '停止', '已请求停止操作')

    def log(self, msg: str):
        """Append a timestamped message to the output terminal."""
        ts = datetime.now().isoformat(sep=' ', timespec='seconds')
        try:
            # terminal may not exist in some contexts
            if hasattr(self, 'terminal') and self.terminal is not None:
                self.terminal.append(f"[{ts}] {msg}")
        except Exception:
            pass

    # -------------- 数据 & 绘图 --------------
    def _on_append_data(self, tup):
        self.data.append(tup)
        self._update_plot()

    def _update_plot(self):
        xs = [d[0] for d in self.data if d[0] is not None]
        ys = [d[1] for d in self.data if d[0] is not None]
        if not xs:
            return
        self.line.set_data(xs, ys)
        self.ax.relim()
        self.ax.autoscale_view()
        self.fig.canvas.draw_idle()

    def save_data(self):
        if not self.data:
            return QMessageBox.warning(self, '无数据', '当前没有数据可保存')
        fn, _ = QFileDialog.getSaveFileName(self, '保存 CSV', '', 'CSV (*.csv)')
        if not fn:
            return
        try:
            with open(fn, 'w', newline='') as f:
                csv.writer(f).writerows([['voltage_V', 'current_A', 'timestamp'], *self.data])
            QMessageBox.information(self, '保存', f'数据已保存到 {fn}')
        except Exception as e:
            QMessageBox.critical(self, '保存失败', str(e))

    # -------------- 几何记忆 --------------
    def _load_geometry(self):
        try:
            with open('gui_geometry.json') as f:
                geo = json.load(f)
                self.setGeometry(geo['x'], geo['y'], geo['w'], geo['h'])
        except Exception:
            self.resize(500, 700)

    def closeEvent(self, event):
        try:
            with open('gui_geometry.json', 'w') as f:
                r = self.geometry()
                tmp = r.split('+')
                w, h = map(int, tmp[0].split('x'))
                x, y = map(int, tmp[1:])
                json.dump({'x': x, 'y': y, 'w': w, 'h': h}, f)
        except Exception:
            pass
        event.accept()

    # -------------- 兼容 tk 版启动接口 --------------
    def mainloop(self):
        """模拟 tk 的 mainloop"""
        from PyQt5.QtWidgets import QApplication
        app = QApplication.instance() or QApplication(sys.argv)
        self.show()
        app.exec_()


# -------------- 供 main.py 调用的 run() --------------
def run():
    app = QApplication.instance() or QApplication(sys.argv)
    w = MainWindow()
    w.mainloop()

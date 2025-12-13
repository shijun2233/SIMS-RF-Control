# app_window.py
import sys
import json
import csv
import time
import threading

import numpy as np

from datetime import datetime

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                             QVBoxLayout, QGridLayout, QGroupBox, QLabel,
                             QLineEdit, QPushButton, QMessageBox, QFileDialog,
                             QTextEdit, QSizePolicy, QComboBox)
from PyQt5.QtCore import Qt, pyqtSignal, QObject, pyqtSlot
from PyQt5.QtGui import QFont

import matplotlib
from matplotlib import pyplot as plt

matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# 假设原控制类在同一目录
from TDK_Control import TDKPowerSupply
from Ammeter_Control import KeithleyPicoammeter


# ------------ 线程通信 ------------
class SigEmitter(QObject):
    append_data = pyqtSignal(tuple)
    update_display = pyqtSignal(dict)  # 用于更新电源显示数据


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('能散测量控制界面')
        # 拓宽窗口，保持高度
        self.resize(900, 360)
        self._load_geometry()

        # 设备 & 数据
        self.tdk = None  # 反射栅网
        self.tdk_lens = None  # 透镜电源 (COM21-6)
        self.tdk_fcup = None  # 法拉第杯抑制电源 (COM11-6)
        self.ser21 = None  # 共享串口 COM21
        self.amm = None
        self.data = []
        self._stop_event = threading.Event()
        self._cont_stop_event = threading.Event()
        self._amm_lock = threading.Lock()  # 避免多线程读取皮安表互相干扰

        # 信号
        self.sig = SigEmitter()
        self.sig.append_data.connect(self._on_append_data)
        self.sig.update_display.connect(self._apply_display_data)

        # 构建 UI
        self._build_ui()
        
        # 定时器更新电源实时数据
        from PyQt5.QtCore import QTimer
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self._schedule_display_update)
        self.update_timer.start(1000)  # 每1000ms更新一次
        self._updating_display = False  # 防止重复更新

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_lo = QVBoxLayout(central)

        # 顶部动作条（保存、停止）
        main_lo.addLayout(self._top_actions())

        # 左列：透镜 + 法拉第杯抑制 + 终端；右列：主电源、皮安表、参数、图表
        h_main = QHBoxLayout()
        left_widget = QWidget()
        left_widget.setLayout(self._left_column())
        left_widget.setMaximumWidth(400)
        h_main.addWidget(left_widget, 1)

        right_col = QVBoxLayout()
        right_col.addLayout(self._top_bar())
        right_col.addLayout(self._bottom_bar())

        h_plot = QHBoxLayout()
        self.canvas = self._build_canvas()
        h_plot.addWidget(self.canvas)
        right_col.addLayout(h_plot)
        ctrl_row = QHBoxLayout()
        btn_clear_plot = QPushButton('Clear')
        btn_clear_plot.setStyleSheet('font-size: 14px; padding: 6px 12px;')
        btn_clear_plot.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        btn_clear_plot.clicked.connect(self.clear_plot)
        btn_stop_cont = QPushButton('停止测量')
        btn_stop_cont.setStyleSheet('font-size: 14px; padding: 6px 12px;')
        btn_stop_cont.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        btn_stop_cont.clicked.connect(self.stop_continuous_measure)
        ctrl_row.addWidget(btn_clear_plot)
        ctrl_row.addWidget(btn_stop_cont)
        ctrl_row.setStretch(0, 1)
        ctrl_row.setStretch(1, 1)
        right_col.addLayout(ctrl_row)

        h_main.addLayout(right_col, 3)
        main_lo.addLayout(h_main)

    # ---------- 顶部工具栏 ----------
    def _top_bar(self):
        lo = QHBoxLayout()

        # TDK（主电源）
        g1 = QGroupBox('反射栅网电源')
        g1.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        g1.setStyleSheet("QGroupBox { color: blue; border: 2px solid blue; border-radius: 5px; padding-top: 10px; margin-top: 10px; }")
        grid = QGridLayout(g1)
        self.pwr_port = QLineEdit('COM21')
        self.pwr_addr = QLineEdit('7')
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
        btn_set = QPushButton('设置')
        btn_on = QPushButton('输出 ON')
        btn_off = QPushButton('输出 OFF')
        btn_connect_p.clicked.connect(self.connect_power)
        btn_disconnect_p.clicked.connect(self.disconnect_power)
        btn_set.clicked.connect(self.set_voltage_current)
        btn_on.clicked.connect(lambda: self.set_output(True))
        btn_off.clicked.connect(lambda: self.set_output(False))

        grid.addWidget(btn_connect_p, 2, 0)
        grid.addWidget(btn_disconnect_p, 2, 1)
        grid.addWidget(btn_set, 0, 4, 2, 1)
        grid.addWidget(btn_on, 2, 2)
        grid.addWidget(btn_off, 2, 3)
        
        # 实时电压电流显示
        self.pwr_actual_v = QLabel('电压: 0.000 V')
        self.pwr_actual_i = QLabel('电流: 0.000 A')
        label_style = 'color: #006400; font-weight: bold; font-size: 12px;'
        self.pwr_actual_v.setStyleSheet(label_style)
        self.pwr_actual_i.setStyleSheet(label_style)
        grid.addWidget(self.pwr_actual_v, 3, 0, 1, 2)
        grid.addWidget(self.pwr_actual_i, 3, 2, 1, 2)
        # Keithley
        g2 = QGroupBox('皮安表')
        g2.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        g2.setStyleSheet("QGroupBox { color: blue; border: 2px solid blue; border-radius: 5px; padding-top: 10px; margin-top: 10px; }")
        grid2 = QGridLayout(g2)
        self.amm_port = QLineEdit('COM12')
        grid2.addWidget(QLabel('串口'), 0, 0)
        grid2.addWidget(self.amm_port, 0, 1)
        btn_connect_a = QPushButton('连接安培表')
        btn_disconnect_a = QPushButton('断开安培表')
        btn_prep = QPushButton('准备测量')
        btn_connect_a.clicked.connect(self.connect_amm)
        btn_disconnect_a.clicked.connect(self.disconnect_amm)
        btn_prep.clicked.connect(self.prepare_measure)
        grid2.addWidget(btn_connect_a, 1, 0)
        grid2.addWidget(btn_disconnect_a, 1, 1)
        self.amm_range_combo = QComboBox()
        self.amm_range_combo.addItems(['自动', '2e-9', '2e-8', '2e-7', '2e-6', '2e-5', '2e-4', '2e-3', '2.1e-2'])
        self.amm_range_combo.setCurrentIndex(0)  # 默认自动量程
        grid2.addWidget(QLabel('量程(A)'), 2, 0)
        grid2.addWidget(self.amm_range_combo, 2, 1)
        grid2.addWidget(btn_prep, 2, 2)
        lo.addWidget(g1)
        lo.addWidget(g2)
        return lo

    def _left_column(self):
        """左侧独立列：透镜电源 + 法拉第杯抑制电源"""
        col = QVBoxLayout()

        # 透镜电源（端口21-6）
        g3 = QGroupBox('透镜电源')
        g3.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        g3.setStyleSheet("QGroupBox { color: blue; border: 2px solid blue; border-radius: 5px; padding-top: 10px; margin-top: 10px; }")
        grid3 = QGridLayout(g3)
        self.lens_port = QLineEdit('COM21')
        self.lens_addr = QLineEdit('6')
        self.lens_v = QLineEdit()
        self.lens_i = QLineEdit()
        grid3.addWidget(QLabel('串口'), 0, 0)
        grid3.addWidget(self.lens_port, 0, 1)
        grid3.addWidget(QLabel('地址'), 1, 0)
        grid3.addWidget(self.lens_addr, 1, 1)
        grid3.addWidget(QLabel('电压(V)'), 0, 2)
        grid3.addWidget(self.lens_v, 0, 3)
        grid3.addWidget(QLabel('电流(A)'), 1, 2)
        grid3.addWidget(self.lens_i, 1, 3)
        btn_l_connect = QPushButton('连接')
        btn_l_disconnect = QPushButton('断开')
        btn_l_set = QPushButton('设置')
        btn_l_on = QPushButton('输出 ON')
        btn_l_off = QPushButton('输出 OFF')
        btn_l_connect.clicked.connect(self.connect_lens)
        btn_l_disconnect.clicked.connect(self.disconnect_lens)
        btn_l_set.clicked.connect(self.set_lens_voltage_current)
        btn_l_on.clicked.connect(lambda: self.set_lens_output(True))
        btn_l_off.clicked.connect(lambda: self.set_lens_output(False))
        grid3.addWidget(btn_l_connect, 2, 0)
        grid3.addWidget(btn_l_disconnect, 2, 1)
        grid3.addWidget(btn_l_set, 0, 4, 2, 1)
        grid3.addWidget(btn_l_on, 2, 2)
        grid3.addWidget(btn_l_off, 2, 3)
        
        # 实时电压电流显示
        self.lens_actual_v = QLabel('电压: 0.000 V')
        self.lens_actual_i = QLabel('电流: 0.000 A')
        label_style = 'color: #006400; font-weight: bold; font-size: 11px;'
        self.lens_actual_v.setStyleSheet(label_style)
        self.lens_actual_i.setStyleSheet(label_style)
        grid3.addWidget(self.lens_actual_v, 3, 0, 1, 2)
        grid3.addWidget(self.lens_actual_i, 3, 2, 1, 2)

        # 法拉第杯抑制电源（端口11-6）
        g4 = QGroupBox('抑制电源 ')
        g4.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        g4.setStyleSheet("QGroupBox { color: blue; border: 2px solid blue; border-radius: 5px; padding-top: 10px; margin-top: 10px; }")
        grid4 = QGridLayout(g4)
        self.fcup_port = QLineEdit('COM11')
        self.fcup_addr = QLineEdit('6')
        self.fcup_v = QLineEdit()
        self.fcup_i = QLineEdit()
        grid4.addWidget(QLabel('串口'), 0, 0)
        grid4.addWidget(self.fcup_port, 0, 1)
        grid4.addWidget(QLabel('地址'), 1, 0)
        grid4.addWidget(self.fcup_addr, 1, 1)
        grid4.addWidget(QLabel('电压(V)'), 0, 2)
        grid4.addWidget(self.fcup_v, 0, 3)
        grid4.addWidget(QLabel('电流(A)'), 1, 2)
        grid4.addWidget(self.fcup_i, 1, 3)
        btn_f_connect = QPushButton('连接')
        btn_f_disconnect = QPushButton('断开')
        btn_f_set = QPushButton('设置')
        btn_f_on = QPushButton('输出 ON')
        btn_f_off = QPushButton('输出 OFF')
        btn_f_connect.clicked.connect(self.connect_fcup)
        btn_f_disconnect.clicked.connect(self.disconnect_fcup)
        btn_f_set.clicked.connect(self.set_fcup_voltage_current)
        btn_f_on.clicked.connect(lambda: self.set_fcup_output(True))
        btn_f_off.clicked.connect(lambda: self.set_fcup_output(False))
        grid4.addWidget(btn_f_connect, 2, 0)
        grid4.addWidget(btn_f_disconnect, 2, 1)
        grid4.addWidget(btn_f_set, 0, 4, 2, 1)
        grid4.addWidget(btn_f_on, 2, 2)
        grid4.addWidget(btn_f_off, 2, 3)
        
        # 实时电压电流显示
        self.fcup_actual_v = QLabel('电压: 0.000 V')
        self.fcup_actual_i = QLabel('电流: 0.000 A')
        label_style = 'color: #006400; font-weight: bold; font-size: 11px;'
        self.fcup_actual_v.setStyleSheet(label_style)
        self.fcup_actual_i.setStyleSheet(label_style)
        grid4.addWidget(self.fcup_actual_v, 3, 0, 1, 2)
        grid4.addWidget(self.fcup_actual_i, 3, 2, 1, 2)

        col.addWidget(g3, 0)
        col.addWidget(g4, 0)
        # 输出终端放在左列底部，占满剩余高度
        self.terminal = QTextEdit()
        self.terminal.setReadOnly(True)
        self.terminal.setPlaceholderText('输出终端')
        self.terminal.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        col.addWidget(self.terminal, 1)
        return col

    def _top_actions(self):
        lo = QHBoxLayout()
        btn_save = QPushButton('保存数据')
        btn_save.clicked.connect(self.save_data)
        lo.addWidget(btn_save)
        hint = QLabel('提示: 清空图片也会清空数据')
        hint.setStyleSheet('color: #666; font-size: 12px;')
        lo.addWidget(hint)
        lo.addStretch()
        return lo

    # ---------- 底部参数 ----------
    def _bottom_bar(self):
        lo = QHBoxLayout()

        # 步进（去掉起始电压，默认从当前设置/实际电压开始）
        g3 = QGroupBox('能散测量参数设置')
        g3.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        g3.setStyleSheet("QGroupBox { color: blue; border: 2px solid blue; border-radius: 5px; padding-top: 10px; margin-top: 10px; }")
        grid3 = QGridLayout(g3)
        self.stop_v, self.step_v, self.step_time = [QLineEdit() for _ in range(3)]
        self.step_time.setText('1')
        labels = ('终止 V', '步长 V', '每步时间(s)')
        for i, (lab, w) in enumerate(zip(labels, (self.stop_v, self.step_v, self.step_time))):
            grid3.addWidget(QLabel(lab), 0, i * 2)
            grid3.addWidget(w, 0, i * 2 + 1)
        btn_go = QPushButton('开始能散测量')
        btn_go.clicked.connect(self.start_step_and_measure)
        grid3.addWidget(btn_go, 1, 0, 1, 6)
        lo.addWidget(g3)

        # 连续测量
        g4 = QGroupBox('连续测量参数')
        g4.setFont(QFont('Microsoft YaHei', 11, QFont.Bold))
        g4.setStyleSheet("QGroupBox { color: blue; border: 2px solid blue; border-radius: 5px; padding-top: 10px; margin-top: 10px; }")
        grid4 = QGridLayout(g4)
        self.measure_steps = QLineEdit('10')
        self.measure_interval = QLineEdit('1')
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
        self.ax.set_xlabel('Measure point')
        self.ax.set_ylabel('Current (A)')
        self.ax.ticklabel_format(style='scientific', scilimits=(0,0), axis='y', useOffset=False)
        self.line, = self.ax.plot([], [], '-o', markersize=4)
        return FigureCanvas(self.fig)

    # ---------------- 原函数名完全一致 ----------------
    def connect_power(self):
        port = self.pwr_port.text().strip()
        addr = int(self.pwr_addr.text().strip() or '1')
        if not port:
            return QMessageBox.warning(self, '提示', '请填写电源串口')
        ser = None
        try:
            import serial
            if self.ser21 and getattr(self.ser21, 'is_open', False) and self.ser21.port == port:
                ser = self.ser21
            else:
                ser = serial.Serial(port, baudrate=9600, timeout=0.5)
                # 仅在 COM21 共享
                if port.lower() == 'com21':
                    self.ser21 = ser
            self.tdk = TDKPowerSupply(addr, ser)
        except Exception as e:
            return QMessageBox.critical(self, '串口打开失败', str(e))
        ok = False
        try:
            ok = self.tdk.test_communication()
        except Exception:
            pass
        QMessageBox.information(self, '连接电源', '通信成功' if ok else '通信可能失败，请检查')
        self.log(f'连接主电源 port={port}, addr={addr}, ok={ok}')
        # 同口连锁：自动准备透镜电源实例
        if port.lower() == 'com21' and self.tdk_lens is None:
            try:
                lens_addr = int(self.lens_addr.text().strip() or '6')
                self.tdk_lens = TDKPowerSupply(lens_addr, ser or self.ser21)
                self.log(f'连锁自动创建透镜电源 addr={lens_addr}')
            except Exception:
                pass

    def disconnect_power(self):
        if self.tdk:
            # 若透镜仍在使用共享串口，则仅释放引用
            if self.tdk_lens:
                self.tdk = None
            else:
                try:
                    self.tdk.disconnect()
                except Exception:
                    pass
                self.tdk = None
                try:
                    if self.ser21:
                        self.ser21.close()
                except Exception:
                    pass
                self.ser21 = None
            QMessageBox.information(self, '断开', '电源已断开')
            self.log('主电源已断开')

    def connect_amm(self):
        port = self.amm_port.text().strip()
        if not port:
            return QMessageBox.warning(self, '提示', '请填写安培表串口')
        self.amm = KeithleyPicoammeter(port)
        ok = self.amm.connect()
        QMessageBox.information(self, '连接安培表', '连接成功' if ok else '连接失败')
        self.log(f'连接皮安表 port={port}, ok={ok}')
        if ok:
            # 成功重连后清空数据与图像
            self.clear_plot()

    def disconnect_amm(self):
        if self.amm:
            self.amm.disconnect()
            self.amm = None
            QMessageBox.information(self, '断开', '安培表已断开')
            self.log('皮安表已断开')

    def set_voltage_current(self):
        if not self.tdk:
            return QMessageBox.warning(self, '未连接', '请先连接电源')
        try:
            v = float(self.voltage_entry.text())
            i = float(self.current_entry.text())
        except Exception:
            return QMessageBox.critical(self, '错误', '无效电压或电流值')
        self.tdk.set_voltage(v)
        self.tdk.set_current(i)
        self.log(f'主电源设定 电压{v}V 电流{i}A')

    def set_output(self, state: bool):
        if not self.tdk:
            return QMessageBox.warning(self, '未连接', '请先连接电源')
        self.tdk.set_output(state)
        self.log(f'主电源输出 {"ON" if state else "OFF"}')

    # -------- 透镜电源 (21-6) --------
    def connect_lens(self):
        port = self.lens_port.text().strip()
        addr = int(self.lens_addr.text().strip() or '6')
        if not port:
            return QMessageBox.warning(self, '提示', '请填写透镜电源串口')
        ser = None
        try:
            import serial
            if self.ser21 and getattr(self.ser21, 'is_open', False) and self.ser21.port == port:
                ser = self.ser21
            else:
                ser = serial.Serial(port, baudrate=9600, timeout=0.5)
                if port.lower() == 'com21':
                    self.ser21 = ser
            self.tdk_lens = TDKPowerSupply(addr, ser)
        except Exception as e:
            return QMessageBox.critical(self, '串口打开失败', str(e))
        ok = False
        try:
            ok = self.tdk_lens.test_communication()
        except Exception:
            pass
        QMessageBox.information(self, '连接透镜电源', '通信成功' if ok else '通信可能失败，请检查')
        self.log(f'连接透镜电源 port={port}, addr={addr}, ok={ok}')
        # 同口连锁：自动准备主电源实例
        if port.lower() == 'com21' and self.tdk is None:
            try:
                main_addr = int(self.pwr_addr.text().strip() or '7')
                self.tdk = TDKPowerSupply(main_addr, ser or self.ser21)
                self.log(f'连锁自动创建主电源 addr={main_addr}')
            except Exception:
                pass

    def disconnect_lens(self):
        if self.tdk_lens:
            if self.tdk:
                self.tdk_lens = None
            else:
                try:
                    self.tdk_lens.disconnect()
                except Exception:
                    pass
                self.tdk_lens = None
                try:
                    if self.ser21:
                        self.ser21.close()
                except Exception:
                    pass
                self.ser21 = None
            QMessageBox.information(self, '断开', '透镜电源已断开')
            self.log('透镜电源已断开')

    def set_lens_voltage_current(self):
        if not self.tdk_lens:
            return QMessageBox.warning(self, '未连接', '请先连接透镜电源')
        try:
            v = float(self.lens_v.text())
            i = float(self.lens_i.text())
        except Exception:
            return QMessageBox.critical(self, '错误', '无效电压或电流值')
        self.tdk_lens.set_voltage(v)
        self.tdk_lens.set_current(i)
        self.log(f'透镜电源设定 电压{v}V 电流{i}A')

    def set_lens_output(self, state: bool):
        if not self.tdk_lens:
            return QMessageBox.warning(self, '未连接', '请先连接透镜电源')
        self.tdk_lens.set_output(state)
        self.log(f'透镜电源输出 {"ON" if state else "OFF"}')

    # -------- 法拉第杯抑制电源 (11-6) --------
    def connect_fcup(self):
        port = self.fcup_port.text().strip()
        addr = int(self.fcup_addr.text().strip() or '6')
        if not port:
            return QMessageBox.warning(self, '提示', '请填写抑制电源串口')
        try:
            import serial
            ser = serial.Serial(port, baudrate=9600, timeout=0.5)
        except Exception as e:
            return QMessageBox.critical(self, '串口打开失败', str(e))
        self.tdk_fcup = TDKPowerSupply(addr, ser)
        ok = self.tdk_fcup.test_communication()
        QMessageBox.information(self, '连接抑制电源', '通信成功' if ok else '通信可能失败，请检查')
        self.log(f'连接抑制电源 port={port}, addr={addr}, ok={ok}')

    def disconnect_fcup(self):
        if self.tdk_fcup:
            self.tdk_fcup.disconnect()
            self.tdk_fcup = None
            QMessageBox.information(self, '断开', '抑制电源已断开')
            self.log('抑制电源已断开')

    def set_fcup_voltage_current(self):
        if not self.tdk_fcup:
            return QMessageBox.warning(self, '未连接', '请先连接抑制电源')
        try:
            v = float(self.fcup_v.text())
            i = float(self.fcup_i.text())
        except Exception:
            return QMessageBox.critical(self, '错误', '无效电压或电流值')
        self.tdk_fcup.set_voltage(v)
        self.tdk_fcup.set_current(i)
        self.log(f'抑制电源设定 电压{v}V 电流{i}A')

    def set_fcup_output(self, state: bool):
        if not self.tdk_fcup:
            return QMessageBox.warning(self, '未连接', '请先连接抑制电源')
        self.tdk_fcup.set_output(state)
        self.log(f'抑制电源输出 {"ON" if state else "OFF"}')

    def apply_range(self):
        # 选择量程仅记录选择，不在此下发命令；命令在 prepare_measure 中统一发送
        return

    def prepare_measure(self):
        if not self.amm:
            return QMessageBox.warning(self, '未连接', '请先连接安培表')
        rng_text = self.amm_range_combo.currentText().strip()
        if rng_text == '自动':
            cmds = ["*RST", "SYST:ACH ON", "RANG 2e-9", "INIT", "SYST:ZCOR:ACQ",
                    "SYST:ZCOR ON","RANG:AUTO ON", "SYST:ZCH OFF"]
        else:
            cmds = ["*RST", "SYST:ACH ON", f"RANG {rng_text}",  "INIT", "SYST:ZCOR:ACQ",
                    "SYST:ZCOR ON", "SYST:ZCH OFF"]
        for c in cmds:
            self.amm.send_command(c)
            time.sleep(0.05)
        QMessageBox.information(self, '准备', '已发送准备测量命令')
        self.log(f'prepare_measure: sent preparation commands, range={rng_text}')

    def single_measure(self):
        if not self.amm:
            return QMessageBox.warning(self, '未连接', '请先连接安培表')
        with self._amm_lock:
            val = self.amm.measure_current()
        if val is None:
            return QMessageBox.critical(self, '测量失败', '未能读取电流')
        volt = self.tdk.get_actual_voltage() if self.tdk else None
        self.sig.append_data.emit((volt, val, datetime.now().isoformat()))
        self.log(f'单次测量完成: I={val}')

    def start_measure(self):
        if not self.amm:
            return QMessageBox.warning(self, '未连接', '请先连接安培表')
        try:
            steps = int(self.measure_steps.text())
            interval = float(self.measure_interval.text())
        except Exception:
            return QMessageBox.critical(self, '错误', '请填写有效的步数与间隔')
        self._cont_stop_event.clear()
        self.log(f'开始连续测量 steps={steps}, interval={interval}s')
        threading.Thread(target=self._measure_loop, args=(steps, interval), daemon=True).start()

    def _measure_loop(self, steps, interval):
        for _ in range(steps):
            if self._cont_stop_event.is_set() or self._stop_event.is_set():
                break
            with self._amm_lock:
                val = self.amm.measure_current()
            volt = self.tdk.get_actual_voltage() if self.tdk else None
            self.sig.append_data.emit((volt, val, datetime.now().isoformat()))
            self.log(f'连续测量: V={volt} I={val}')
            # 使用Event.wait()替代循环sleep，响应更快
            if self._cont_stop_event.wait(interval) or self._stop_event.is_set():
                break

    def stop_continuous_measure(self):
        self._cont_stop_event.set()
        QMessageBox.information(self, '停止', '已停止当前皮安表连续测量')
        self.log('已请求停止连续测量')

    def start_step_and_measure(self):
        if not (self.tdk and self.amm):
            return QMessageBox.warning(self, '未连接', '请先连接电源与安培表')
        try:
            # 起始电压去除：从当前电压开始（优先实际输出，其次输入框）
            start = self.tdk.get_actual_voltage()
            if start is None:
                start = float(self.voltage_entry.text() or '0')
            stop, step, step_time = map(float, (self.stop_v.text(), self.step_v.text(), self.step_time.text()))
        except Exception:
            return QMessageBox.critical(self, '错误', '请填写有效的步进参数')
        if (stop - start) * step < 0:
            return QMessageBox.critical(self, '错误', '步长方向与起止不匹配')
        self._stop_event.clear()
        self.log(f'开始阶梯输出 start={start}, stop={stop}, step={step}, hold={step_time}s')
        threading.Thread(target=self._step_and_measure_thread,
                         args=(start, stop, step, step_time), daemon=True).start()

    def _step_and_measure_thread(self, start, stop, step, step_time):
        volt, ascending = start, step > 0
        eps = 1e-12
        while True:
            # 超界直接退出
            if (ascending and volt > stop + eps) or (not ascending and volt < stop - eps):
                break
            if self._stop_event.is_set():
                break
            step_start = time.perf_counter()
            self.tdk.set_voltage(volt)
            time.sleep(0.05)  # 简短稳压时间计入总步长
            # 在步长的 1/2 处再测一次，使用Event.wait优化
            mid_delay = step_time / 2.0
            if self._stop_event.wait(mid_delay):
                break
            if self._stop_event.is_set():
                break
            mid_cur = None
            mid_attempts = 0
            while mid_cur is None and not self._stop_event.is_set():
                with self._amm_lock:
                    mid_cur = self.amm.measure_current()
                if mid_cur is None:
                    mid_attempts += 1
                    if mid_attempts >= 10:
                        self.log(f'阶梯中点测量失败, V={volt}, 尝试{mid_attempts}次')
                        break
                    time.sleep(0.1)
            if mid_cur is None:
                break
            self.sig.append_data.emit((volt, mid_cur, datetime.now().isoformat()))
            self.log(f'阶梯测量: V={volt} I={mid_cur}')
            # 使用Event.wait等待剩余时间
            elapsed = time.perf_counter() - step_start
            remaining = step_time - elapsed
            if remaining > 0:
                if self._stop_event.wait(remaining):
                    break
            # 计算下一步，避免跨过终止值，确保终止值也执行
            next_vol = volt + step
            if ascending and next_vol > stop:
                next_vol = stop
            if (not ascending) and next_vol < stop:
                next_vol = stop
            # 若当前已经是终止值，完成后退出
            if abs(volt - stop) <= eps:
                break
            volt = next_vol

    def stop_operations(self):
        self._stop_event.set()
        QMessageBox.information(self, '停止', '已请求停止操作')
        self.log('已请求停止所有操作')

    def log(self, msg: str):
        """Append a timestamped message to the output terminal."""
        if not hasattr(self, 'terminal') or self.terminal is None:
            return
        ts = datetime.now().isoformat(sep=' ', timespec='seconds')
        self.terminal.append(f"[{ts}] {msg}")

    # -------------- 更新电源显示 --------------
    def _schedule_display_update(self):
        """调度后台更新，避免阻塞UI"""
        if self._updating_display:
            return  # 上次更新未完成，跳过
        self._updating_display = True
        threading.Thread(target=self._update_power_displays_background, daemon=True).start()
    
    def _update_power_displays_background(self):
        """后台线程读取电源数据"""
        data = {}
        try:
            if self.tdk:
                v = self.tdk.get_actual_voltage()
                i = self.tdk.get_actual_current()
                if v is not None and i is not None:
                    data['pwr'] = (v, i)
        except Exception:
            pass
        
        try:
            if self.tdk_lens:
                v = self.tdk_lens.get_actual_voltage()
                i = self.tdk_lens.get_actual_current()
                if v is not None and i is not None:
                    data['lens'] = (v, i)
        except Exception:
            pass
        
        try:
            if self.tdk_fcup:
                v = self.tdk_fcup.get_actual_voltage()
                i = self.tdk_fcup.get_actual_current()
                if v is not None and i is not None:
                    data['fcup'] = (v, i)
        except Exception:
            pass
        
        # 使用信号更新UI（线程安全）
        if data:
            self.sig.update_display.emit(data)
        self._updating_display = False
    
    @pyqtSlot(dict)
    def _apply_display_data(self, data):
        """在主线程中更新显示（槽函数）"""
        if 'pwr' in data:
            v, i = data['pwr']
            self.pwr_actual_v.setText(f'电压: {v:.3f} V')
            self.pwr_actual_i.setText(f'电流: {i:.3f} A')
        if 'lens' in data:
            v, i = data['lens']
            self.lens_actual_v.setText(f'电压: {v:.3f} V')
            self.lens_actual_i.setText(f'电流: {i:.3f} A')
        if 'fcup' in data:
            v, i = data['fcup']
            self.fcup_actual_v.setText(f'电压: {v:.3f} V')
            self.fcup_actual_i.setText(f'电流: {i:.3f} A')

    # -------------- 数据 & 绘图 --------------
    def _on_append_data(self, tup):
        self.data.append(tup)
        try:
            v, cur, ts = tup
            self.log(f"测量数据: time={ts}, V={v}, I={cur}")
        except Exception:
            pass
        self._update_plot()


    def _update_plot(self):
        # x 轴为测量点序号，y 轴为电流
        indices, currents = [], []
        for v, cur, ts in self.data:
            if cur is None:
                continue
            indices.append(len(indices) + 1)
            currents.append(cur)
        if not indices:
            return
        self.line.set_data(indices, currents)
        self.ax.relim()
        self.ax.autoscale_view(tight=True)
        self.fig.canvas.draw_idle()


    def clear_plot(self):
        """清理图与数据"""
        self.data = []
        self.line.set_data([], [])
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
            rows = []
            for v, cur, ts in self.data:
                rows.append([ts, f'{v}V' if v is not None else '', cur])
            with open(fn, 'w', newline='') as f:
                csv.writer(f).writerows([['time', 'voltage_V','current_A'], *rows])
            QMessageBox.information(self, '保存', f'数据已保存到 {fn}')
            self.log(f'保存数据 -> {fn}')
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

import threading
import time
import csv
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from TDK_Control import TDKPowerSupply
from Ammeter_Control import KeithleyPicoammeter


class MainWindow(tk.Tk):
	def __init__(self):
		super().__init__()
		self.title('TDK 电源与安培表联合控制')
		self.geometry('1000x700')

		# 设备实例
		self.tdk = None
		self.amm = None

		# 数据保存
		self.data = []  # list of (voltage, current, timestamp)

		# UI 布局
		self._build_ui()

		# 测量线程控制
		self._stop_event = threading.Event()

	def _build_ui(self):
		# 左侧控制区
		left = ttk.Frame(self)
		left.pack(side=tk.LEFT, fill=tk.Y, padx=8, pady=8)

		# 电源控制区
		pwr_frame = ttk.LabelFrame(left, text='TDK 电源')
		pwr_frame.pack(fill=tk.X, pady=4)

		ttk.Label(pwr_frame, text='串口:').grid(row=0, column=0, sticky='w')
		self.pwr_port = ttk.Entry(pwr_frame)
		self.pwr_port.grid(row=0, column=1)
		# 默认电源串口
		self.pwr_port.insert(0, 'COM11')

		ttk.Label(pwr_frame, text='地址(NSEL):').grid(row=1, column=0, sticky='w')
		self.pwr_addr = ttk.Entry(pwr_frame)
		self.pwr_addr.insert(0, '3')
		self.pwr_addr.grid(row=1, column=1)

		ttk.Button(pwr_frame, text='连接电源', command=self.connect_power).grid(row=2, column=0, pady=4)
		ttk.Button(pwr_frame, text='断开电源', command=self.disconnect_power).grid(row=2, column=1, pady=4)

		ttk.Label(pwr_frame, text='电压(V):').grid(row=3, column=0, sticky='w')
		self.voltage_entry = ttk.Entry(pwr_frame)
		self.voltage_entry.grid(row=3, column=1)
		ttk.Button(pwr_frame, text='设置电压', command=self.set_voltage).grid(row=3, column=2)

		ttk.Label(pwr_frame, text='电流(A):').grid(row=4, column=0, sticky='w')
		self.current_entry = ttk.Entry(pwr_frame)
		self.current_entry.grid(row=4, column=1)
		ttk.Button(pwr_frame, text='设置电流', command=self.set_current).grid(row=4, column=2)

		ttk.Button(pwr_frame, text='输出 ON', command=lambda: self.set_output(True)).grid(row=5, column=0)
		ttk.Button(pwr_frame, text='输出 OFF', command=lambda: self.set_output(False)).grid(row=5, column=1)

		# 步进设置
		step_frame = ttk.LabelFrame(left, text='步进输出设置')
		step_frame.pack(fill=tk.X, pady=4)
		ttk.Label(step_frame, text='起始V').grid(row=0, column=0)
		self.start_v = ttk.Entry(step_frame, width=8)
		self.start_v.grid(row=0, column=1)
		ttk.Label(step_frame, text='终止V').grid(row=0, column=2)
		self.stop_v = ttk.Entry(step_frame, width=8)
		self.stop_v.grid(row=0, column=3)
		ttk.Label(step_frame, text='步长V').grid(row=1, column=0)
		self.step_v = ttk.Entry(step_frame, width=8)
		self.step_v.grid(row=1, column=1)
		ttk.Label(step_frame, text='每步时间(s)').grid(row=1, column=2)
		self.step_time = ttk.Entry(step_frame, width=8)
		self.step_time.insert(0, '0.2')
		self.step_time.grid(row=1, column=3)

		ttk.Button(step_frame, text='开始阶梯输出并测量', command=self.start_step_and_measure).grid(row=2, column=0, columnspan=2, pady=6)
		ttk.Button(step_frame, text='停止', command=self.stop_operations).grid(row=2, column=2, columnspan=2)

		# 安培表控制区
		amm_frame = ttk.LabelFrame(left, text='安培表 (Keithley)')
		amm_frame.pack(fill=tk.X, pady=4)
		ttk.Label(amm_frame, text='串口:').grid(row=0, column=0)
		self.amm_port = ttk.Entry(amm_frame)
		self.amm_port.grid(row=0, column=1)
		# 默认安培表串口
		self.amm_port.insert(0, 'COM12')

		ttk.Button(amm_frame, text='连接安培表', command=self.connect_amm).grid(row=1, column=0)
		ttk.Button(amm_frame, text='断开安培表', command=self.disconnect_amm).grid(row=1, column=1)

		ttk.Button(amm_frame, text='选择电源测量', command=self.select_source_measure).grid(row=2, column=0)
		ttk.Button(amm_frame, text='准备测量', command=self.prepare_measure).grid(row=2, column=1)

		ttk.Label(amm_frame, text='测量步数:').grid(row=3, column=0)
		self.measure_steps = ttk.Entry(amm_frame, width=6)
		self.measure_steps.insert(0, '10')
		self.measure_steps.grid(row=3, column=1)

		ttk.Label(amm_frame, text='测量间隔(s):').grid(row=4, column=0)
		self.measure_interval = ttk.Entry(amm_frame, width=6)
		self.measure_interval.insert(0, '0.2')
		self.measure_interval.grid(row=4, column=1)

		ttk.Button(amm_frame, text='开始测量', command=self.start_measure).grid(row=5, column=0)
		ttk.Button(amm_frame, text='单次测量', command=self.single_measure).grid(row=5, column=1)

		ttk.Button(left, text='保存数据', command=self.save_data).pack(pady=8)

		# 右侧绘图区
		right = ttk.Frame(self)
		right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)

		self.fig = Figure(figsize=(6,4), dpi=100)
		self.ax = self.fig.add_subplot(111)
		self.ax.set_xlabel('Voltage (V)')
		self.ax.set_ylabel('Current (A)')
		self.line, = self.ax.plot([], [], '-o')

		self.canvas = FigureCanvasTkAgg(self.fig, master=right)
		self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

	# ---------------- Device connection methods ----------------
	def connect_power(self):
		port = self.pwr_port.get().strip()
		addr = int(self.pwr_addr.get().strip() or '1')
		if not port:
			messagebox.showwarning('提示','请填写电源串口')
			return
		try:
			ser = __import__('serial').Serial(port, baudrate=9600, timeout=0.5)
		except Exception as e:
			messagebox.showerror('串口打开失败', str(e))
			return
		self.tdk = TDKPowerSupply(addr, ser)
		ok = self.tdk.test_communication()
		messagebox.showinfo('连接电源', '通信成功' if ok else '通信可能失败，请检查')

	def disconnect_power(self):
		if self.tdk:
			self.tdk.disconnect()
			self.tdk = None
			messagebox.showinfo('断开', '电源已断开')

	def connect_amm(self):
		port = self.amm_port.get().strip()
		if not port:
			messagebox.showwarning('提示','请填写安培表串口')
			return
		self.amm = KeithleyPicoammeter(port)
		ok = self.amm.connect()
		messagebox.showinfo('连接安培表', '连接成功' if ok else '连接失败')

	def disconnect_amm(self):
		if self.amm:
			self.amm.disconnect()
			self.amm = None
			messagebox.showinfo('断开', '安培表已断开')

	# ---------------- Power control ----------------
	def set_voltage(self):
		if not self.tdk:
			messagebox.showwarning('未连接', '请先连接电源')
			return
		try:
			v = float(self.voltage_entry.get())
		except Exception:
			messagebox.showerror('错误', '无效电压值')
			return
		self.tdk.set_voltage(v)

	def set_current(self):
		if not self.tdk:
			messagebox.showwarning('未连接', '请先连接电源')
			return
		try:
			i = float(self.current_entry.get())
		except Exception:
			messagebox.showerror('错误', '无效电流值')
			return
		self.tdk.set_current(i)

	def set_output(self, state: bool):
		if not self.tdk:
			messagebox.showwarning('未连接', '请先连接电源')
			return
		self.tdk.set_output(state)

	# ---------------- Ammeter specialized controls ----------------
	def select_source_measure(self):
		# Placeholder: depending on instrument, selection may require wiring or commands
		messagebox.showinfo('提示', '请选择电源测量（硬件接线）')

	def prepare_measure(self):
		if not self.amm:
			messagebox.showwarning('未连接', '请先连接安培表')
			return
		# 按用户要求发送一系列命令
		cmds = ["*RST", "SYST:ACH ON", "RANG 2e-9", "INIT", "SYST:ZCOR:ACQ", "SYST:ZCOR ON", "RANG:AUTO ON", "SYST:ZCH OFF"]
		for c in cmds:
			self.amm.send_command(c)
			time.sleep(0.05)
		messagebox.showinfo('准备', '已发送准备测量命令')

	def single_measure(self):
		if not self.amm:
			messagebox.showwarning('未连接', '请先连接安培表')
			return
		val = self.amm.measure_current()
		if val is None:
			messagebox.showerror('测量失败', '未能读取电流')
			return
		timestamp = datetime.now().isoformat()
		# 如果有电源实例，尝试读取实际电压
		volt = None
		if self.tdk:
			volt = self.tdk.get_actual_voltage()
		self.data.append((volt, val, timestamp))
		self._update_plot()
		messagebox.showinfo('测量结果', f'电流: {val} A')

	def start_measure(self):
		if not self.amm:
			messagebox.showwarning('未连接', '请先连接安培表')
			return
		try:
			steps = int(self.measure_steps.get())
			interval = float(self.measure_interval.get())
		except Exception:
			messagebox.showerror('错误', '请填写有效的步数与间隔')
			return

		self._stop_event.clear()
		t = threading.Thread(target=self._measure_loop, args=(steps, interval), daemon=True)
		t.start()

	def _measure_loop(self, steps, interval):
		for i in range(steps):
			if self._stop_event.is_set():
				break
			val = self.amm.measure_current()
			timestamp = datetime.now().isoformat()
			volt = None
			if self.tdk:
				volt = self.tdk.get_actual_voltage()
			self.data.append((volt, val, timestamp))
			self._update_plot()
			time.sleep(interval)

	def start_step_and_measure(self):
		if not self.tdk or not self.amm:
			messagebox.showwarning('未连接', '请先连接电源与安培表')
			return
		try:
			start = float(self.start_v.get())
			stop = float(self.stop_v.get())
			step = float(self.step_v.get())
			step_time = float(self.step_time.get())
		except Exception:
			messagebox.showerror('错误', '请填写有效的步进参数')
			return

		self._stop_event.clear()
		t = threading.Thread(target=self._step_and_measure_thread, args=(start, stop, step, step_time), daemon=True)
		t.start()

	def _step_and_measure_thread(self, start, stop, step, step_time):
		# 构建数列，支持增减
		if step == 0:
			return
		if (stop - start) * step < 0:
			messagebox.showerror('错误', '步长方向与起止不匹配')
			return
		volt = start
		ascending = step > 0
		while True:
			if (ascending and volt > stop) or (not ascending and volt < stop):
				break
			if self._stop_event.is_set():
				break
			# 设置电压
			self.tdk.set_voltage(volt)
			# wait small settle time
			time.sleep(0.2)
			# 测量电流
			cur = self.amm.measure_current()
			timestamp = datetime.now().isoformat()
			self.data.append((volt, cur, timestamp))
			self._update_plot()
			# 等待该步时长
			elapsed = 0.0
			while elapsed < step_time:
				if self._stop_event.is_set():
					break
				time.sleep(0.05)
				elapsed += 0.05
			volt += step

	def stop_operations(self):
		self._stop_event.set()
		messagebox.showinfo('停止', '已请求停止操作')

	# ---------------- Plot & Save ----------------
	def _update_plot(self):
		# 更新 matplotlib 图表，按电压排序
		xs = [d[0] for d in self.data if d[0] is not None]
		ys = [d[1] for d in self.data if d[0] is not None]
		if not xs:
			return
		try:
			self.line.set_data(xs, ys)
			self.ax.relim()
			self.ax.autoscale_view()
			self.canvas.draw_idle()
		except Exception:
			pass

	def save_data(self):
		if not self.data:
			messagebox.showwarning('无数据', '当前没有数据可保存')
			return
		fn = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV', '*.csv')])
		if not fn:
			return
		try:
			with open(fn, 'w', newline='') as f:
				w = csv.writer(f)
				w.writerow(['voltage_V', 'current_A', 'timestamp'])
				for row in self.data:
					w.writerow(row)
			messagebox.showinfo('保存', f'数据已保存到 {fn}')
		except Exception as e:
			messagebox.showerror('保存失败', str(e))


def run():
	app = MainWindow()
	app.mainloop()


if __name__ == '__main__':
	run()


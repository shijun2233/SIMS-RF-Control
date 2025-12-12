

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial
import socket
import threading
import queue

import serial.tools.list_ports

class SCPI_Controller:
    """
    Handles communication with the instrument via Serial or TCP/IP.
    """
    def __init__(self):
        self.instrument = None
        self.connection_type = None

    def get_available_ports(self):
        """Returns a list of available serial ports."""
        return [port.device for port in serial.tools.list_ports.comports()]

    def connect_serial(self, port, baudrate, timeout=1):
        """Establishes a serial connection."""
        try:
            self.instrument = serial.Serial(port, int(baudrate), timeout=timeout)
            self.connection_type = 'serial'
            return f"Successfully connected to {port} at {baudrate} baud."
        except Exception as e:
            self.instrument = None
            return f"Error connecting to {port}: {e}"

    def connect_tcp(self, host, port, timeout=2):
        """Establishes a TCP/IP socket connection."""
        try:
            self.instrument = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.instrument.settimeout(timeout)
            self.instrument.connect((host, int(port)))
            self.connection_type = 'tcp'
            return f"Successfully connected to {host}:{port}."
        except Exception as e:
            self.instrument = None
            return f"Error connecting to {host}:{port}: {e}"

    def disconnect(self):
        """Closes the connection."""
        if self.instrument:
            try:
                self.instrument.close()
                message = "Disconnected."
            except Exception as e:
                message = f"Error during disconnection: {e}"
            self.instrument = None
            self.connection_type = None
            return message
        return "Already disconnected."

    def send_command(self, command):
        """Sends a command to the instrument."""
        if not self.instrument:
            raise ConnectionError("Not connected to any instrument.")
        
        # Ensure command is bytes and ends with a newline
        if not command.endswith('\n'):
            command += '\n'
        
        try:
            if self.connection_type == 'serial':
                self.instrument.write(command.encode('ascii'))
            elif self.connection_type == 'tcp':
                self.instrument.sendall(command.encode('ascii'))
        except Exception as e:
            raise IOError(f"Failed to send command: {e}")

    def read_response(self, buffer_size=4096):
        """Reads a response from the instrument."""
        if not self.instrument:
            raise ConnectionError("Not connected to any instrument.")
        
        try:
            if self.connection_type == 'serial':
                response = self.instrument.readline().decode('ascii').strip()
            elif self.connection_type == 'tcp':
                response = self.instrument.recv(buffer_size).decode('ascii').strip()
            return response
        except socket.timeout:
            return "Timeout: No response from instrument."
        except Exception as e:
            raise IOError(f"Failed to read response: {e}")

    def query(self, command):
        """Sends a command and reads the response."""
        self.send_command(command)
        return self.read_response()


class App(tk.Tk):
    """
    GUI for the SCPI Controller.
    """
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.title("SCPI Command Sender")
        self.geometry("600x550") # Increased height for new fields

        self.response_queue = queue.Queue()
        self.after(100, self.process_queue)

        self.create_widgets()
        self.update_serial_ports()

    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Connection Frame ---
        conn_frame = ttk.LabelFrame(main_frame, text="Connection", padding="10")
        conn_frame.pack(fill=tk.X, pady=5)
        conn_frame.grid_columnconfigure(1, weight=1)

        # Connection Type
        self.conn_type_var = tk.StringVar(value="Serial")
        ttk.Label(conn_frame, text="Type:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        serial_rb = ttk.Radiobutton(conn_frame, text="Serial", variable=self.conn_type_var, value="Serial", command=self.toggle_connection_fields)
        serial_rb.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        tcp_rb = ttk.Radiobutton(conn_frame, text="TCP/IP", variable=self.conn_type_var, value="TCP/IP", command=self.toggle_connection_fields)
        tcp_rb.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

        # Serial Port
        self.port_label = ttk.Label(conn_frame, text="Port:")
        self.port_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(conn_frame, textvariable=self.port_var, state="readonly")
        self.port_combo.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW)
        self.refresh_button = ttk.Button(conn_frame, text="Refresh", command=self.update_serial_ports)
        self.refresh_button.grid(row=1, column=3, padx=5, pady=5)

        # Baudrate
        self.baud_label = ttk.Label(conn_frame, text="Baudrate:")
        self.baud_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.baud_var = tk.StringVar(value="9600")
        self.baud_entry = ttk.Entry(conn_frame, textvariable=self.baud_var)
        self.baud_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW)

        # Address Enable
        self.addr_enable_var = tk.BooleanVar(value=False)
        self.addr_enable_check = ttk.Checkbutton(conn_frame, text="启用地址", variable=self.addr_enable_var, command=self.toggle_address_field)
        self.addr_enable_check.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

        # Address Entry
        self.addr_label = ttk.Label(conn_frame, text="地址:")
        self.addr_var = tk.StringVar(value="1")
        self.addr_entry = ttk.Entry(conn_frame, textvariable=self.addr_var)

        # IP Address
        self.ip_label = ttk.Label(conn_frame, text="IP Address:")
        self.ip_var = tk.StringVar(value="192.168.1.100")
        self.ip_entry = ttk.Entry(conn_frame, textvariable=self.ip_var)

        # TCP Port
        self.tcp_port_label = ttk.Label(conn_frame, text="Port:")
        self.tcp_port_var = tk.StringVar(value="5025")
        self.tcp_port_entry = ttk.Entry(conn_frame, textvariable=self.tcp_port_var)

        # Connect/Disconnect Buttons
        self.connect_button = ttk.Button(conn_frame, text="Connect", command=self.connect)
        self.connect_button.grid(row=4, column=1, padx=5, pady=10, sticky=tk.E)
        self.disconnect_button = ttk.Button(conn_frame, text="Disconnect", command=self.disconnect, state=tk.DISABLED)
        self.disconnect_button.grid(row=4, column=2, padx=5, pady=10, sticky=tk.W)

        # --- Command Frame ---
        cmd_frame = ttk.LabelFrame(main_frame, text="SCPI Command", padding="10")
        cmd_frame.pack(fill=tk.X, pady=5)
        cmd_frame.grid_columnconfigure(0, weight=1)

        self.cmd_var = tk.StringVar(value="*IDN?")
        self.cmd_entry = ttk.Entry(cmd_frame, textvariable=self.cmd_var)
        self.cmd_entry.grid(row=0, column=0, padx=(0, 5), pady=5, sticky=tk.EW)
        self.send_button = ttk.Button(cmd_frame, text="Send", command=self.send_command, state=tk.DISABLED)
        self.send_button.grid(row=0, column=1, padx=5, pady=5)
        self.query_button = ttk.Button(cmd_frame, text="Query", command=self.query_command, state=tk.DISABLED)
        self.query_button.grid(row=0, column=2, padx=5, pady=5)

        # --- Log/Response Frame ---
        log_frame = ttk.LabelFrame(main_frame, text="Log and Response", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state='disabled')

        self.toggle_connection_fields()
        self.toggle_address_field()

    def toggle_address_field(self):
        """Show/hide address field based on the checkbox."""
        if self.addr_enable_var.get() and self.conn_type_var.get() == "Serial":
            self.addr_label.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
            self.addr_entry.grid(row=3, column=2, padx=5, pady=5, sticky=tk.EW)
        else:
            self.addr_label.grid_remove()
            self.addr_entry.grid_remove()

    def toggle_connection_fields(self):
        """Show/hide fields based on connection type."""
        conn_type = self.conn_type_var.get()
        if conn_type == "Serial":
            self.port_label.grid()
            self.port_combo.grid()
            self.refresh_button.grid()
            self.baud_label.grid()
            self.baud_entry.grid()
            self.addr_enable_check.grid()
            
            self.ip_label.grid_remove()
            self.ip_entry.grid_remove()
            self.tcp_port_label.grid_remove()
            self.tcp_port_entry.grid_remove()
        else: # TCP/IP
            self.port_label.grid_remove()
            self.port_combo.grid_remove()
            self.refresh_button.grid_remove()
            self.baud_label.grid_remove()
            self.baud_entry.grid_remove()
            self.addr_enable_check.grid_remove()

            self.ip_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
            self.ip_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW)
            self.tcp_port_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
            self.tcp_port_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW)
        
        self.toggle_address_field()

    def update_serial_ports(self):
        """Refresh the list of available serial ports."""
        ports = self.controller.get_available_ports()
        self.port_combo['values'] = ports
        if ports:
            self.port_var.set(ports[0])

    def log(self, message):
        """Append a message to the log text area."""
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.configure(state='disabled')
        self.log_text.see(tk.END)

    def connect(self):
        conn_type = self.conn_type_var.get()
        status = ""
        if conn_type == "Serial":
            port = self.port_var.get()
            baud = self.baud_var.get()
            if not port or not baud:
                messagebox.showerror("Error", "Port and Baudrate cannot be empty.")
                return
            status = self.controller.connect_serial(port, baud)
        else: # TCP/IP
            host = self.ip_var.get()
            port = self.tcp_port_var.get()
            if not host or not port:
                messagebox.showerror("Error", "IP Address and Port cannot be empty.")
                return
            status = self.controller.connect_tcp(host, port)
        
        self.log(status)
        if self.controller.instrument:
            self.set_ui_state(connected=True)
        else:
            messagebox.showerror("Connection Failed", status)

    def disconnect(self):
        status = self.controller.disconnect()
        self.log(status)
        self.set_ui_state(connected=False)

    def set_ui_state(self, connected):
        """Enable/disable UI elements based on connection status."""
        state = tk.NORMAL if connected else tk.DISABLED
        self.send_button.config(state=state)
        self.query_button.config(state=state)
        self.disconnect_button.config(state=state)

        conn_state = tk.DISABLED if connected else tk.NORMAL
        self.connect_button.config(state=conn_state)
        for child in self.connect_button.master.winfo_children():
            if isinstance(child, (ttk.Entry, ttk.Combobox, ttk.Radiobutton, ttk.Button, ttk.Checkbutton)) and child not in [self.connect_button, self.disconnect_button]:
                 child.config(state=conn_state)

    def run_in_thread(self, target, *args):
        """Run a function in a separate thread to avoid blocking the GUI."""
        thread = threading.Thread(target=target, args=args, daemon=True)
        thread.start()

    def get_full_command(self):
        """Constructs the full command, prepending address if enabled."""
        command = self.cmd_var.get()
        if self.conn_type_var.get() == "Serial" and self.addr_enable_var.get():
            address = self.addr_var.get()
            if address:
                return f"{address};{command}"
        return command

    def send_command(self):
        command = self.get_full_command()
        if not command:
            return
        self.log(f"SEND: {command}")
        self.run_in_thread(self._send_worker, command)

    def _send_worker(self, command):
        try:
            self.controller.send_command(command)
        except Exception as e:
            self.response_queue.put(f"ERROR: {e}")

    def query_command(self):
        command = self.get_full_command()
        if not command:
            return
        self.log(f"QUERY: {command}")
        self.run_in_thread(self._query_worker, command)

    def _query_worker(self, command):
        try:
            response = self.controller.query(command)
            self.response_queue.put(f"RECV: {response}")
        except Exception as e:
            self.response_queue.put(f"ERROR: {e}")

    def process_queue(self):
        """Process messages from the worker threads."""
        try:
            while True:
                message = self.response_queue.get_nowait()
                self.log(message)
        except queue.Empty:
            pass
        self.after(100, self.process_queue)

    def on_closing(self):
        """Handle window closing event."""
        if self.controller.instrument:
            self.disconnect()
        self.destroy()

if __name__ == "__main__":
    scpi_controller = SCPI_Controller()
    app = App(scpi_controller)
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
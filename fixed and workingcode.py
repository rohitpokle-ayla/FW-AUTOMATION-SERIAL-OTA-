import serial
import serial.tools.list_ports
import threading
import time
import os
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import scrolledtext, simpledialog, messagebox

# Auto create folders with absolute paths
script_dir = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(script_dir, "logs"), exist_ok=True)
os.makedirs(os.path.join(script_dir, "reports"), exist_ok=True)
os.makedirs(os.path.join(script_dir, "ota_logs"), exist_ok=True)

# Global variables
log_lines = []
fixtures = []
events = []
missing_fixtures = []
stop_thread = False
selected_port = None
serial_thread = None
stop_logging = False
device_id = None
ota_triggered = False
last_update_time = time.time()
ser = None  # Global serial connection


class OTASystem:
    def __init__(self):
        self.driver = None
        self.log_filename = None
        self._initialize_driver()
        
    def _initialize_driver(self):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        driver_path = r"C:\Users\RohitPokle\Downloads\chromedriver-win64 (1)\chromedriver-win64\chromedriver.exe"
        self.driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)
        
    def _initialize_logfile(self):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_dir = os.path.join(script_dir, "ota_logs")
        os.makedirs(log_dir, exist_ok=True)
        self.log_filename = os.path.join(log_dir, f"{timestamp}_ota_log.txt")
        
    def start_serial_logging(self, port):
        """Start logging serial output during OTA process"""
        self._initialize_logfile()
        ser = serial.Serial(port, 115200, timeout=1)
        
        with open(self.log_filename, "w") as logfile:
            while not stop_logging:
                if ser.in_waiting:
                    line = ser.readline().decode(errors="ignore").strip()
                    if line:
                        ts_line = f"{datetime.now().strftime('%H:%M:%S')} - {line}"
                        logfile.write(ts_line + "\n")
                        logfile.flush()
                        
        ser.close()
        
    def perform_ota_update(self, device_id):
        """Perform the OTA update process"""
        try:
            # Step 1: Navigate to login page
            self.driver.get("https://dashboard.aylanetworks.com/sessions/new")
            time.sleep(5)

            try:
                no_button = self.driver.find_element(By.XPATH, "//button[@type='button' and contains(text(), 'No')]")
                no_button.click()
                print("Dismissed modal popup by clicking 'No'")
            except:
                print("No popup appeared.")

            # Step 2: Enter Email
            email_input = self.driver.find_element(By.NAME, "email")
            email_input.send_keys("rohit.rpokle+volt@gmail.com")

            # Step 3: Enter Password
            password_input = self.driver.find_element(By.NAME, "password")
            password_input.send_keys("Ayla1234")

            # Step 4: Click Login
            login_button = self.driver.find_element(By.CLASS_NAME, "log-in-button")
            login_button.click()
            time.sleep(10)

            # Step 6: Click on Search Label
            search_label = self.driver.find_element(By.CSS_SELECTOR, "label[data-modal-title='Search']")
            search_label.click()
            time.sleep(5)

            # Step 8: Enter DSN value 
            dsn_input = self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='DSN']")
            dsn_input.send_keys(device_id)

            # Step 9: Click Search
            search_button = self.driver.find_element(By.XPATH, "//button[text()='SEARCH']")
            search_button.click()
            time.sleep(5)

            try:
                # Wait until the device ID row is visible
                status_td = WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//td[contains(text(), '{device_id}')]/preceding-sibling::td")
                    )
                )
                print("✅ Device status element found.")

                # Check status color (example)
                status_color = status_td.value_of_css_property("background-color")
                print("Status Color:", status_color)

                if "red" in status_color or "rgb(255" in status_color:
                    print("❌ Device is OFFLINE. Please turn it ON.")
                    self.driver.save_screenshot("device_offline.png")
                    return False
                else:
                    print("✅ Device is ONLINE. Continuing...")

            except Exception as e:
                print(f"❌ Error finding device status: {e}")
                self.driver.save_screenshot("element_not_found.png")
                return False

            # Click on the DSN cell
            dsn_cell = self.driver.find_element(By.XPATH, f"//td[contains(text(), '{device_id}')]")
            dsn_cell.click()

            # Scroll to and click the Host MCU Images tab using ng-click
            try:
                host_mcu_tab = WebDriverWait(self.driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(@ng-click, 'showDetail') and contains(@ng-click, 'host-MCU-images')]"))
                )
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", host_mcu_tab)
                WebDriverWait(self.driver, 15).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@ng-click, 'showDetail') and contains(@ng-click, 'host-MCU-images')]"))).click()
                print("✅ Scrolled to and clicked Host MCU Images tab.")
            except Exception as e:
                print("❌ Could not scroll to or click Host MCU Images tab.")
                print("Message:", e)
                self.driver.save_screenshot("host_mcu_tab_scroll_click_failed.png")
                return False

            # Click on OTA image icon
            try:
                ota_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "span.aylaicon.aylaicon-add-ota"))
                )
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", ota_button)
                ota_button.click()
                print("✅ OTA image button clicked.")
            except Exception as e:
                print("❌ Could not find or click OTA image button.")
                print("Message:", e)
                self.driver.save_screenshot("ota_button_not_found.png")
                self.driver.quit()
                exit()

            # Click the Accept button in modal
            try:
                accept_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Accept') and @ng-click='modal.confirm.onAccept()']"))
                )
                accept_btn.click()
                print("✅ Accepted OTA update.")
            except Exception as e:
                print("❌ Could not click Accept in modal.")
                print("Message:", e)
                self.driver.save_screenshot("accept_modal_failed.png")
                return False

            # Go to 'Commands' tab
            try:
                commands_tab = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@ng-click, 'showDetail') and contains(text(), 'Commands')]"))
                )
                commands_tab.click()
                print("✅ Navigated to Commands tab.")
            except Exception as e:
                print("❌ Could not click Commands tab.")
                print("Message:", e)
                self.driver.save_screenshot("commands_tab_not_found.png")
                return False

            # Wait until the last command becomes 'true'
            try:
                WebDriverWait(self.driver, 60).until(
                    lambda d: d.find_elements(By.XPATH, "//td[contains(text(), 'true')]")[-1].text.strip().lower() == 'true'
                )
                print("✅ Last command is TRUE.")
                return True
            except Exception as e:
                print("❌ Timeout waiting for last command to be true.")
                print("Message:", e)
                self.driver.save_screenshot("command_not_true.png")
                return False

        except Exception as e:
            print(f"❌ Unexpected error during OTA: {e}")
            return False


class DeviceValidationSystem:
    def __init__(self, log_text_widget=None, app=None):
        self.serial_port = None
        self.baud_rate = 115200
        self.log_filename = None
        self.report_filename = None
        self.start_time = time.time()
        self.log_text_widget = log_text_widget
        self.app = app  # Reference to main application
        self._initialize_filenames()
        self.serial_connection = None

    def _initialize_filenames(self):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_dir = os.path.join(script_dir, "logs")
        report_dir = os.path.join(script_dir, "reports")
        os.makedirs(log_dir, exist_ok=True)
        os.makedirs(report_dir, exist_ok=True)

        self.log_filename = os.path.join(log_dir, f"{timestamp}_log.txt")
        self.report_filename = os.path.join(report_dir, f"{timestamp}_summary.xlsx")

    def _extract_info_from_line(self, line):
        global device_id
        ts = datetime.now().strftime("%H:%M:%S")
        print(f"[{ts}] {line}")
        
        if self.log_text_widget:
            self.log_text_widget.insert(tk.END, f"[{ts}] {line}\n")
            self.log_text_widget.see(tk.END)
        
        keyword_events = {
            "In read_I2C_Slave EEPROM Process success": "✅ I2C EEPROM read success",
            "fault log snapshots saved": "⚠️ Fault log snapshots saved",
            "net_status associating": "🔄 WiFi associating",
            "join succeeded": "✅ WiFi join succeeded",
            "net_status disconnected": "❌ WiFi disconnected",
            "net_status connected": "✅ WiFi connected",
            "time set by SNTP": "⏰ Time synced",
            "MESH INIT": "🧩 Mesh initialized",
            "Vendor Key Binded": "🔑 Vendor Key Binded",
            "BLOB Client Key Binded": "🧱 BLOB Client Key Binded",
            "DFU Client Key Binded": "📦 DFU Client Key Binded",
            "volt_ota_save_done": "✅ OTA Save Done",
            "Final OTA status 200 reported": "✅ OTA Status 200 Reported",
            "MQTT client abort error": "❌ MQTT Abort / Connection Lost",
            "Network status: disconnected": "❌ Network Disconnected",
            "Log client disabled": "⚠️ Log Client Disabled – Device Reset",
            "RESET: RTC_SW_CPU_RST": "🔁 Software Reset Triggered",
            "Bootloader & partition table info": "🧰 Bootloader & Partition Info",
            "Firmware boot:": "🚀 Firmware Boot",
        }

        for keyword, message in keyword_events.items():
            if keyword in line:
                events.append((ts, message))
                break
                
        # DSN ID - Only extract DSN but don't stop logging
        if "get DSN" in line:
            device_id = line.split()[-1]
            events.append((ts, f"🔢 DSN: {device_id}"))
            # Enable OTA button when DSN is detected
            if self.app:
                self.app.after(0, self.app.enable_ota_button)

        # MQTT Host
        elif "DNS: host" in line:
            host = line.strip().split("host")[-1].strip()
            events.append((ts, f"☁️ MQTT Host: {host}"))

        # Module Name
        elif "module name" in line:
            module = line.strip().split("module name")[-1].strip()
            events.append((ts, f"📦 Module: {module}"))

        # MAC-ID and Addr fixture
        elif "device_v4 Addr" in line:
            mac = re.search(r"MAC-ID:([\da-fA-F]+)", line)
            addr = re.search(r"Addr:\s*(\d+)", line)
            if mac and addr:
                fixtures.append((addr.group(1), mac.group(1)))

        # IP assignment
        elif "ada_client_ip_up" in line and "IP" in line:
            match = re.search(r"IP\s([\d\.]+)", line)
            if match:
                events.append((ts, f"🌐 IP assigned: {match.group(1)}"))

        # OTA progress
        ota_match = re.search(r"OTA (\d+)% saved", line)
        if ota_match:
            percent = ota_match.group(1)
            events.append((ts, f"📦 OTA {percent}% saved"))

        # OTA fetch done
        if "OTA fetch done" in line:
            match = re.search(r"OTA fetch done: ([\d,]+) bytes", line)
            if match:
                size = match.group(1)
                events.append((ts, f"📥 OTA Fetch Done: {size} bytes"))
            else:
                events.append((ts, f"📥 OTA Fetch Done"))

    def send_serial_command(self, command):
        """Send a command to the serial device"""
        if not self.serial_connection or not self.serial_connection.is_open:
            print("❌ Serial connection not established")
            return False
        
        try:
            # Ensure command ends with newline
            if not command.endswith('\n'):
                command += '\n'
                
            self.serial_connection.write(command.encode())
            print(f"📤 Sent command: {command.strip()}")
            return True
        except Exception as e:
            print(f"❌ Error sending command: {e}")
            return False

    def serial_reader(self):
        global stop_thread, last_update_time, ser
        try:
            self.serial_connection = ser = serial.Serial(
                self.serial_port, 
                self.baud_rate, 
                timeout=1
            )
            
            print(f"📡 Listening to {self.serial_port}... Logs saving to {self.log_filename}")

            with open(self.log_filename, "w") as logfile:
                while not stop_thread:
                    try:
                        if ser.in_waiting:
                            line = ser.readline().decode(errors="ignore").strip()
                            if line:
                                ts_line = f"{datetime.now().strftime('%H:%M:%S')} - {line}"
                                log_lines.append(ts_line)
                                logfile.write(ts_line + "\n")
                                logfile.flush()
                                self._extract_info_from_line(line)

                                if time.time() - last_update_time >= 60:
                                    print("\n🔄 System Updated - Logs and reports being saved")
                                    last_update_time = time.time()

                    except Exception as e:
                        print(f"⚠️ Error reading serial: {e}")
                        break
        except Exception as e:
            print(f"❌ Error initializing serial: {e}")
        finally:
            if ser and ser.is_open:
                ser.close()

    def generate_report(self):
        try:
            wb = Workbook()

            # Timeline Sheet
            ws1 = wb.active
            ws1.title = "Event Timeline"
            ws1.append(["Timestamp", "Event"])
            for row in events:
                ws1.append(row)

            # Fixtures Sheet
            ws2 = wb.create_sheet("Fixtures")
            ws2.append(["Address", "MAC-ID"])
            for item in fixtures:
                ws2.append(item)

            # Missing Fixtures Sheet
            ws3 = wb.create_sheet("Missing Fixtures")
            ws3.append(["Missing Fixture IDs"])
            for item in missing_fixtures:
                ws3.append([item])

            # Bold headers
            for sheet in wb.worksheets:
                for cell in sheet["1:1"]:
                    cell.font = Font(bold=True)

            wb.save(self.report_filename)
            print(f"\n✅ Report generated at: {os.path.abspath(self.report_filename)}")
            return self.report_filename

        except Exception as e:
            print(f"❌ Error generating report: {e}")
            return None

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Device Validation and OTA System")
        self.geometry("1000x700")
        
        self.validation_system = None
        self.ota_system = OTASystem()
        self.serial_thread = None
        self.command_frame = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main container
        main_frame = tk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel (controls)
        left_panel = tk.Frame(main_frame, width=200)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        
        # Serial Port Selection
        port_frame = tk.LabelFrame(left_panel, text="Serial Port Configuration", padx=5, pady=5)
        port_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.port_var = tk.StringVar()
        self.port_menu = tk.OptionMenu(self, self.port_var, "")
        self.port_menu.pack()
        self.refresh_ports()
        
        tk.Label(port_frame, text="Select Port:").pack(side=tk.TOP, anchor=tk.W)
        
        self.refresh_btn = tk.Button(
            port_frame, 
            text="Refresh Ports", 
            command=self.refresh_ports
        )
        self.refresh_btn.pack(fill=tk.X, padx=5, pady=2)
        
        # Command Frame
        self.command_frame = tk.LabelFrame(left_panel, text="Device Commands", padx=5, pady=5)
        self.command_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.reset_btn = tk.Button(
            self.command_frame, 
            text="Send Reset", 
            command=lambda: self.send_command("reset"),
            state=tk.DISABLED
        )
        self.reset_btn.pack(fill=tk.X, padx=5, pady=2)
        
        self.custom_cmd_btn = tk.Button(
            self.command_frame, 
            text="Custom Command", 
            command=self.send_custom_command,
            state=tk.DISABLED
        )
        self.custom_cmd_btn.pack(fill=tk.X, padx=5, pady=2)
        
        # Control Buttons
        control_frame = tk.LabelFrame(left_panel, text="System Controls", padx=5, pady=5)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.start_btn = tk.Button(
            control_frame, 
            text="Start Monitoring", 
            command=self.start_monitoring
        )
        self.start_btn.pack(fill=tk.X, padx=5, pady=2)
        
        self.stop_btn = tk.Button(
            control_frame, 
            text="Stop Monitoring", 
            command=self.stop_monitoring, 
            state=tk.DISABLED
        )
        self.stop_btn.pack(fill=tk.X, padx=5, pady=2)
        
        self.ota_btn = tk.Button(
            control_frame, 
            text="Start OTA", 
            command=self.start_ota, 
            state=tk.DISABLED
        )
        self.ota_btn.pack(fill=tk.X, padx=5, pady=2)
        
        self.report_btn = tk.Button(
            control_frame, 
            text="Generate Report", 
            command=self.generate_report
        )
        self.report_btn.pack(fill=tk.X, padx=5, pady=2)
        
        # Log Display
        log_frame = tk.LabelFrame(main_frame, text="Serial Log", padx=5, pady=5)
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            wrap=tk.WORD, 
            width=100, 
            height=30
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
    def refresh_ports(self):
        ports = serial.tools.list_ports.comports()
        port_devices = [port.device for port in ports]
        if hasattr(self, 'port_menu'): 
            menu = self.port_menu["menu"]
            menu.delete(0, "end")
        
        if port_devices:
            self.port_var.set(port_devices[0])
            for port in port_devices:
                menu.add_command(
                    label=port, 
                    command=lambda p=port: self.port_var.set(p)
                )
        else:
            self.port_var.set("No ports found")
            menu.add_command(
                label="No ports found", 
                command=tk._setit(self.port_var, "No ports found")
            )
            
    def enable_ota_button(self):
        """Enable the OTA button when DSN is detected"""
        self.ota_btn.config(state=tk.NORMAL)
            
    def start_monitoring(self):
        global stop_thread, device_id, log_lines, fixtures, events, missing_fixtures
        stop_thread = False
        device_id = None
        log_lines = []
        fixtures = []
        events = []
        missing_fixtures = []
        
        if self.port_var.get() == "No ports found":
            self.log_text.insert(tk.END, "❌ No serial ports available!\n")
            return
        
        self.validation_system = DeviceValidationSystem(log_text_widget=self.log_text, app=self)
        self.validation_system.serial_port = self.port_var.get()
        
        self.serial_thread = threading.Thread(
            target=self.validation_system.serial_reader,
            daemon=True
        )
        self.serial_thread.start()
        
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.reset_btn.config(state=tk.NORMAL)
        self.custom_cmd_btn.config(state=tk.NORMAL)
        
        self.log_text.insert(tk.END, f"📡 Started monitoring on {self.port_var.get()}\n")
        
    def stop_monitoring(self):
        global stop_thread
        stop_thread = True
        
        if self.serial_thread and self.serial_thread.is_alive():
            self.serial_thread.join(timeout=2)
            
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.reset_btn.config(state=tk.DISABLED)
        self.custom_cmd_btn.config(state=tk.DISABLED)
        self.ota_btn.config(state=tk.DISABLED)
        
        self.log_text.insert(tk.END, "⏹️ Stopped monitoring\n")
            
    def send_command(self, command):
        if not self.validation_system or not self.validation_system.serial_connection:
            self.log_text.insert(tk.END, "❌ Not connected to serial port\n")
            return
            
        success = self.validation_system.send_serial_command(command)
        if success:
            self.log_text.insert(tk.END, f"📤 Sent command: {command}\n")
        else:
            self.log_text.insert(tk.END, f"❌ Failed to send command: {command}\n")
            
    def send_custom_command(self):
        command = simpledialog.askstring(
            "Custom Command", 
            "Enter command to send:",
            parent=self
        )
        
        if command:
            self.send_command(command)
            
    def start_ota(self):
        global ota_triggered, stop_logging, device_id
        
        if not device_id:
            self.log_text.insert(tk.END, "❌ No device ID detected for OTA\n")
            return
            
        response = messagebox.askyesno(
            "Confirm OTA", 
            f"Start OTA for device {device_id}?",
            parent=self
        )
        
        if not response:
            return
            
        ota_triggered = True
        
        # Don't set stop_logging to False here - let the main logging continue
        self.log_text.insert(tk.END, f"\n🚀 Starting OTA update for device {device_id}...\n")
        
        # Disable OTA button during process
        self.ota_btn.config(state=tk.DISABLED)
        
        # Start OTA in a separate thread
        ota_thread = threading.Thread(
            target=self._perform_ota_update_wrapper,
            args=(device_id,),
            daemon=True
        )
        ota_thread.start()
        
    def _perform_ota_update_wrapper(self, device_id):
        try:
            # Start OTA-specific logging in parallel (without affecting main logging)
            ota_log_thread = threading.Thread(
                target=self.ota_system.start_serial_logging,
                args=(self.port_var.get(),),
                daemon=True
            )
            ota_log_thread.start()
            
            # Perform OTA
            success = self.ota_system.perform_ota_update(device_id)
            
            if success:
                self.safe_log_insert("\n🎉 OTA update completed successfully!\n")
                
                # Wait for 5 minutes after OTA
                self.safe_log_insert("⏳ Waiting 5 minutes for post-OTA inspection...\n")
                time.sleep(300)
                self.safe_log_insert("✅ Post-OTA wait completed\n")
            else:
                self.safe_log_insert("\n❌ OTA update failed\n")
                
        except Exception as e:
            self.safe_log_insert(f"\n⚠️ OTA error: {str(e)}\n")
        finally:
            # Re-enable OTA button if device is still connected
            if not stop_thread and device_id:
                self.after(0, lambda: self.ota_btn.config(state=tk.NORMAL))
            
    def safe_log_insert(self, text):
        """Thread-safe way to insert text into the log widget"""
        self.after(0, lambda: self.log_text.insert(tk.END, text))
        self.after(0, lambda: self.log_text.see(tk.END))
        
    def generate_report(self):
        if self.validation_system:
            report_path = self.validation_system.generate_report()
            if report_path:
                self.log_text.insert(
                    tk.END, 
                    f"\n✅ Report generated at: {os.path.abspath(report_path)}\n"
                )
        else:
            self.log_text.insert(
                tk.END, 
                "\n❌ No data available to generate report. Start monitoring first.\n"
            )


def main():
    app = Application()
    app.mainloop()


if __name__ == "__main__":
    main()
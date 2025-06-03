import time
import math
import tkinter as tk
from tkinter import ttk, messagebox
import cv2
import numpy as np
import serial
import serial.tools.list_ports
import seabreeze.spectrometers as sb
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageTk
import threading
import queue
import logging
import os
import pandas as pd
from datetime import datetime

try:
    import wmi
except ImportError:
    wmi = None
    logging.warning("wmi module not found, using psutil fallback")

try:
    import win32com.client
except ImportError as e:
    logging.warning(f"Failed to import win32com.client: {str(e)}")
    win32com = None

import psutil
import ctypes
import subprocess

# Setup logging
logging.basicConfig(
    filename="moon_scanner.log",
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(message)s",
)

# Constants
FULL_MOON_SIZE_ARCSEC = 40
SCAN_DURATION = 10
MINUTES_PER_SAMPLE = 10
TOTAL_SAMPLES = int(SCAN_DURATION * 60 / MINUTES_PER_SAMPLE)
LATITUDE = 40.0
DEFAULT_SPEC_INTEGRATION_MS = 100
DEFAULT_CAM_EXPOSURE_MS = 33.3  # Approx 30 FPS
MAX_QUEUE_SIZE = 10

# Data directories
DATA_DIR = "scan_data"
IMAGE_DIR = os.path.join(DATA_DIR, "images")
SPECTRA_DIR = os.path.join(DATA_DIR, "spectra")

# Timeout decorator
def timeout(seconds):
    def decorator(func):
        def wrapper(*args, **kwargs):
            result = [None]
            exception = [None]
            event = threading.Event()

            def target():
                try:
                    result[0] = func(*args, **kwargs)
                except Exception as e:
                    exception[0] = e
                finally:
                    event.set()

            thread = threading.Thread(target=target)
            thread.daemon = True
            thread.start()
            if not event.wait(seconds):
                raise TimeoutError(f"Function {func.__name__} timed out after {seconds} seconds")
            if exception[0]:
                raise exception[0]
            return result[0]

        return wrapper

    return decorator

class MoonScannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Moon Scanner")
        self.root.geometry("1000x550")
        self.scanning = False
        self.recording = False
        self.camera_running = False
        self.spectrograph_running = False
        self.capture_running = False
        self.data_queue = queue.Queue(maxsize=MAX_QUEUE_SIZE)
        self.data_queue_lock = threading.Lock()
        self.dark_spectrum = None
        self.running = True
        self.camera_index = 0
        self.available_cameras = []
        self.serial_port = "COM3"
        self.mount_initialized = False
        self.mount = None
        self.ascom_telescope = None
        self.ascom_driver = "ASCOM.Celestron.Telescope"
        self.camera = None
        self.spectrograph = None
        self.motor_speed = 0
        self.ascom_slew_speed = 5
        self.spec_integration_ms = DEFAULT_SPEC_INTEGRATION_MS
        self.cam_exposure_ms = DEFAULT_CAM_EXPOSURE_MS
        self.sample_count = 0
        self.scan_angle = 0.0  # degrees
        self.scan_step = 10.0  # arcsec
        self.scan_speed = 1.0  # arcsec/sec
        self.use_camera_exposure = tk.BooleanVar(value=True)  # Toggle for camera exposure

        # Create data directories
        os.makedirs(IMAGE_DIR, exist_ok=True)
        os.makedirs(SPECTRA_DIR, exist_ok=True)
        try:
            with open(os.path.join(DATA_DIR, "test.txt"), "w") as f:
                f.write("test")
            os.remove(os.path.join(DATA_DIR, "test.txt"))
            logging.info(f"Write access confirmed for {DATA_DIR}")
        except Exception as e:
            logging.error(f"No write access to {DATA_DIR}: {str(e)}")
            messagebox.showerror(
                "Error", f"No write access to {DATA_DIR}: {str(e)}. Run as admin or change folder permissions."
            )

        self.check_admin_mode()
        self.setup_gui()
        self.setup_plot()
        self.detect_cameras()

    def check_admin_mode(self):
        if not ctypes.windll.shell32.IsUserAnAdmin():
            logging.warning("Not running as administrator")
            messagebox.showwarning(
                "Admin Mode", "Please run as administrator for full access to camera and COM ports."
            )

    def detect_cameras(self):
        """Detect available cameras."""
        self.available_cameras = []
        index = 0
        while index < 10:  # Check up to 10 camera indices
            cap = cv2.VideoCapture(index, cv2.CAP_DSHOW)
            if cap.isOpened():
                self.available_cameras.append(str(index))
                cap.release()
            index += 1
        if self.available_cameras:
            self.camera_menu['menu'].delete(0, 'end')
            for cam in self.available_cameras:
                self.camera_menu['menu'].add_command(
                    label=cam, command=lambda value=cam: self.camera_index_var.set(value)
                )
            self.camera_index_var.set(self.available_cameras[0])
            logging.info(f"Detected cameras: {self.available_cameras}")
        else:
            logging.warning("No cameras detected")
            messagebox.showwarning("Warning", "No cameras detected. Check connections and drivers.")

    def setup_gui(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=0)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=0)

        # Top frame
        top_frame = ttk.Frame(self.root)
        top_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        top_frame.grid_columnconfigure(0, weight=1)
        top_frame.grid_columnconfigure(1, weight=1)

        # Camera frame
        camera_frame = ttk.LabelFrame(top_frame, text="Camera Feed", padding=5)
        camera_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        ttk.Label(camera_frame, text="Camera:").grid(row=0, column=0, sticky="w")
        self.camera_index_var = tk.StringVar(value="0")
        self.camera_menu = ttk.OptionMenu(camera_frame, self.camera_index_var, "0", *self.available_cameras)
        self.camera_menu.grid(row=0, column=1, pady=2, sticky="w")
        ttk.Button(camera_frame, text="Test Camera", command=self.test_camera).grid(row=0, column=2, pady=2)
        ttk.Label(camera_frame, text="Exposure (ms):").grid(row=1, column=0, sticky="w")
        self.cam_exposure_entry = ttk.Entry(camera_frame, width=8)
        self.cam_exposure_entry.insert(0, str(DEFAULT_CAM_EXPOSURE_MS))
        self.cam_exposure_entry.grid(row=1, column=1, pady=2, sticky="w")
        ttk.Button(camera_frame, text="Set Exposure", command=self.set_camera_exposure).grid(row=1, column=2, pady=2)
        ttk.Checkbutton(
            camera_frame, text="Use Exposure Time", variable=self.use_camera_exposure
        ).grid(row=2, column=0, columnspan=2, pady=2, sticky="w")
        self.camera_button = ttk.Button(camera_frame, text="Start Camera", command=self.toggle_camera)
        self.camera_button.grid(row=3, column=0, pady=2)
        ttk.Button(camera_frame, text="Retry Camera", command=self.retry_camera).grid(row=3, column=1, pady=2)
        ttk.Button(camera_frame, text="Reset Camera", command=self.reset_camera).grid(row=3, column=2, pady=2)
        self.webcam_label = ttk.Label(camera_frame, text="No camera feed")
        self.webcam_label.grid(row=4, column=0, columnspan=3, sticky="n")

        # Spectrum frame
        spectrum_frame = ttk.LabelFrame(top_frame, text="Spectrum", padding=5)
        spectrum_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")
        ttk.Label(spectrum_frame, text="Integration (ms):").grid(row=0, column=0, sticky="w")
        self.spec_integration_entry = ttk.Entry(spectrum_frame, width=8)
        self.spec_integration_entry.insert(0, str(DEFAULT_SPEC_INTEGRATION_MS))
        self.spec_integration_entry.grid(row=0, column=1, pady=2, sticky="w")
        ttk.Button(spectrum_frame, text="Set Integration", command=self.set_spectrograph_integration).grid(
            row=0, column=2, pady=2
        )
        self.spectrograph_button = ttk.Button(
            spectrum_frame, text="Start Spectrograph", command=self.toggle_spectrograph
        )
        self.spectrograph_button.grid(row=1, column=0, pady=2)
        ttk.Button(spectrum_frame, text="Capture Dark", command=self.capture_dark).grid(
            row=1, column=1, columnspan=2, pady=2
        )
        self.plot_frame = ttk.Frame(spectrum_frame)
        self.plot_frame.grid(row=2, column=0, columnspan=3, sticky="n")

        # Control frame
        control_frame = ttk.LabelFrame(self.root, text="Controls", padding=5)
        control_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")
        control_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8), weight=1)
        self.init_button = ttk.Button(control_frame, text="Initialize Mount", command=self.initialize_mount)
        self.init_button.grid(row=0, column=0, padx=2, pady=2)
        ttk.Button(control_frame, text="Reinitialize Mount", command=self.reinitialize_mount).grid(
            row=0, column=1, padx=2, pady=2
        )
        ttk.Button(control_frame, text="Refresh GUI", command=self.refresh_gui).grid(row=0, column=2, padx=2, pady=2)
        ttk.Button(control_frame, text="Test ASCOM", command=self.test_ascom).grid(row=0, column=3, padx=2, pady=2)
        ttk.Button(control_frame, text="Test Serial", command=self.test_serial).grid(row=0, column=4, padx=2, pady=2)
        ttk.Button(control_frame, text="Force COM3 Reset", command=self.force_com3_reset).grid(row=0, column=5, padx=2, pady=2)
        self.scan_button = ttk.Button(control_frame, text="Start Scan", command=self.start_scan)
        self.scan_button.grid(row=0, column=6, padx=2, pady=2)
        self.stop_scan_button = ttk.Button(control_frame, text="Stop Scan", command=self.stop_scan)
        self.stop_scan_button.grid(row=0, column=7, padx=2, pady=2)
        self.record_button = ttk.Button(control_frame, text="Start Recording", command=self.toggle_record)
        self.record_button.grid(row=0, column=8, padx=2, pady=2)
        # Scan parameters
        ttk.Label(control_frame, text="Scan Angle (deg):").grid(row=1, column=0, sticky="w")
        self.scan_angle_entry = ttk.Entry(control_frame, width=8)
        self.scan_angle_entry.insert(0, "0.2")
        self.scan_angle_entry.grid(row=1, column=1, pady=2, sticky="w")
        ttk.Button(control_frame, text="Set Angle", command=self.set_scan_angle).grid(row=1, column=2, pady=2)
        ttk.Label(control_frame, text="Step Offset (arcsec):").grid(row=2, column=0, sticky="w")
        self.scan_step_entry = ttk.Entry(control_frame, width=8)
        self.scan_step_entry.insert(0, "10.0")
        self.scan_step_entry.grid(row=2, column=1, pady=2, sticky="w")
        ttk.Button(control_frame, text="Set Step", command=self.set_scan_step).grid(row=2, column=2, pady=2)
        ttk.Label(control_frame, text="Scan Speed (arcsec/s):").grid(row=3, column=0, sticky="w")
        self.scan_speed_entry = ttk.Entry(control_frame, width=8)
        self.scan_speed_entry.insert(0, "1.0")
        self.scan_speed_entry.grid(row=3, column=1, pady=2, sticky="w")
        ttk.Button(control_frame, text="Set Speed", command=self.set_scan_speed).grid(row=3, column=2, pady=2)
        self.status_label = ttk.Label(control_frame, text="Status: Idle")
        self.status_label.grid(row=4, column=0, columnspan=9, pady=2, sticky="w")

        # Mount control frame
        self.mount_control_frame = ttk.LabelFrame(self.root, text="Mount Control", padding=5)
        self.mount_control_frame.grid(row=1, column=2, padx=5, pady=5, sticky="ns")
        ttk.Label(self.mount_control_frame, text="Serial Port:").grid(row=0, column=0, sticky="w")
        self.serial_port_var = tk.StringVar(value="COM3")
        self.serial_port_menu = ttk.OptionMenu(
            self.mount_control_frame, self.serial_port_var, "COM3", "COM3", "COM4", "COM5", "COM6"
        )
        self.serial_port_menu.grid(row=0, column=1, pady=2)
        self.mount_north_button = ttk.Button(
            self.mount_control_frame, text="^", command=lambda: self.move_mount_manual("North"), width=3
        )
        self.mount_north_button.grid(row=1, column=1, padx=2, pady=2)
        self.mount_west_button = ttk.Button(
            self.mount_control_frame, text="<-", command=lambda: self.move_mount_manual("West"), width=3
        )
        self.mount_west_button.grid(row=2, column=0, padx=2, pady=2)
        self.mount_south_button = ttk.Button(
            self.mount_control_frame, text="v", command=lambda: self.move_mount_manual("South"), width=3
        )
        self.mount_south_button.grid(row=2, column=1, padx=2, pady=2)
        self.mount_east_button = ttk.Button(
            self.mount_control_frame, text="->", command=lambda: self.move_mount_manual("East"), width=3
        )
        self.mount_east_button.grid(row=2, column=2, padx=2, pady=2)
        self.mount_stop_button = ttk.Button(
            self.mount_control_frame, text="Stop", command=self.stop_mount_manual, width=5
        )
        self.mount_stop_button.grid(row=3, column=1, padx=2, pady=2)
        ttk.Label(self.mount_control_frame, text="Speed (1-9):").grid(row=4, column=0, sticky="w")
        self.slew_speed_var = tk.StringVar(value=str(self.ascom_slew_speed))
        ttk.Entry(self.mount_control_frame, textvariable=self.slew_speed_var, width=5).grid(row=4, column=1, pady=2)
        ttk.Button(self.mount_control_frame, text="Set", command=self.set_slew_speed, width=5).grid(
            row=4, column=2, padx=2, pady=2
        )
        self.update_mount_controls()

    def setup_plot(self):
        self.fig, self.ax = plt.subplots(figsize=(4, 2.5))
        self.line, = self.ax.plot([], [], "b-")
        self.ax.set_xlabel("Wavelength (nm)")
        self.ax.set_ylabel("Intensity")
        self.ax.set_title("Spectrum")
        self.ax.set_xlim(200, 1000)
        self.ax.grid(True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plot_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="n")

    def update_record_button_style(self):
        style = ttk.Style()
        if self.recording:
            style.configure("Record.TButton", background="green", foreground="white")
            self.record_button.configure(style="Record.TButton")
        else:
            style.configure("Record.TButton", background="red", foreground="white")
            self.record_button.configure(style="Record.TButton")

    def update_mount_controls(self):
        state = "normal" if self.mount_initialized and self.ascom_telescope else "disabled"
        self.scan_button.config(state=state)
        self.stop_scan_button.config(state=state)
        self.mount_north_button.config(state=state)
        self.mount_west_button.config(state=state)
        self.mount_south_button.config(state=state)
        self.mount_east_button.config(state=state)
        self.mount_stop_button.config(state=state)
        logging.debug(f"Mount controls updated: state={state}")

    def set_camera_exposure(self):
        if not self.camera:
            messagebox.showerror("Error", "Camera not initialized. Click 'Start Camera'.")
            return
        if not self.use_camera_exposure.get():
            logging.info("Camera exposure time disabled by user")
            self.status_label.config(text="Status: Camera exposure time disabled")
            return
        try:
            exposure_ms = float(self.cam_exposure_entry.get())
            if exposure_ms <= 0:
                raise ValueError("Exposure time must be positive")
            self.cam_exposure_ms = exposure_ms
            exposure_val = exposure_ms / 1000.0  # Approximate seconds
            self.camera.set(cv2.CAP_PROP_EXPOSURE, exposure_val)
            actual_exposure = self.camera.get(cv2.CAP_PROP_EXPOSURE)
            logging.info(f"Set camera exposure to {exposure_ms:.2f}ms (actual: {actual_exposure*1000:.2f}ms)")
            self.status_label.config(text=f"Status: Camera exposure set to {exposure_ms:.1f}ms")
        except ValueError as e:
            logging.error(f"Invalid exposure time: {str(e)}")
            messagebox.showerror("Error", f"Invalid exposure time: {str(e)}")
        except Exception as e:
            logging.error(f"Failed to set exposure: {str(e)}")
            messagebox.showerror("Error", f"Failed to set exposure: {str(e)}. Camera may not support exposure control.")

    def set_scan_angle(self):
        try:
            angle = float(self.scan_angle_entry.get())
            if angle <= 0:
                raise ValueError("Scan angle must be positive")
            self.scan_angle = angle
            logging.info(f"Scan angle set to {angle:.2f}°")
            self.status_label.config(text=f"Status: Scan angle set to {angle:.2f}°")
        except ValueError as e:
            logging.error(f"Invalid scan angle: {str(e)}")
            messagebox.showerror("Error", f"Invalid scan angle: {str(e)}")
        except Exception as e:
            logging.error(f"Error setting scan angle: {str(e)}")
            messagebox.showerror("Error", f"Error setting scan angle: {str(e)}")

    def set_scan_step(self):
        try:
            step = float(self.scan_step_entry.get())
            if step <= 0:
                raise ValueError("Step offset must be positive")
            self.scan_step = step
            logging.info(f"Scan step set to {step:.2f} arcsec")
            self.status_label.config(text=f"Status: Scan step set to {step:.2f} arcsec")
            messagebox.showinfo("Success", "Scan step set successfully")
        except ValueError as e:
            logging.error(f"Invalid scan step: {str(e)}")
            messagebox.showerror("Error", f"Invalid scan step: {str(e)}")
        except Exception as e:
            logging.error(f"Error setting scan step: {str(e)}")
            messagebox.showerror("Error", f"Failed to set scan step: {str(e)}")

    def set_scan_speed(self):
        try:
            speed = float(self.scan_speed_entry.get())
            if speed <= 0:
                raise ValueError("Scan speed must be positive")
            self.scan_speed = speed
            logging.info(f"Scan speed set to {speed:.2f} arcsec/s")
            self.status_label.config(text=f"Status: Scan speed set to {speed:.2f} arcsec/s")
        except ValueError as e:
            logging.error(f"Invalid scan speed: {str(e)}")
            messagebox.showerror("Error", f"Invalid scan speed: {str(e)}")
        except Exception as e:
            logging.error(f"Error setting scan speed: {str(e)}")
            messagebox.showerror("Error", f"Failed to set scan speed: {str(e)}")

    def check_com3_status(self):
        """Check if COM3 is free and log processes using it."""
        try:
            logging.info("Checking COM3 status")
            processes_using_com3 = []
            for proc in psutil.process_iter(['pid', 'name', 'open_files', 'cmdline']):
                try:
                    proc_info = f"PID {proc.pid}: {proc.name()} (cmd: {' '.join(proc.cmdline()) if proc.cmdline() else 'N/A'})"
                    for f in proc.open_files() or []:
                        if 'COM3' in f.path.upper():
                            processes_using_com3.append(proc_info)
                            logging.info(f"COM3 in use by {proc_info}")
                    if proc.name().lower() in ['serialport.exe', 'cpwi.exe', 'skyportal.exe']:
                        processes_using_com3.append(proc_info)
                        logging.info(f"Potential COM3 user: {proc_info}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                except Exception as e:
                    logging.error(f"Error checking process {proc.pid}: {str(e)}")

            try:
                result = subprocess.run(['netstat', '-a', '-o'], capture_output=True, text=True)
                for line in result.stdout.splitlines():
                    if 'COM3' in line.upper():
                        logging.info(f"Netstat: {line}")
            except Exception as e:
                logging.error(f"Netstat check failed: {str(e)}")

            if processes_using_com3:
                return False, f"COM3 in use by: {', '.join(processes_using_com3)}"
            return True, "COM3 is free"
        except Exception as e:
            logging.error(f"COM3 status check failed: {str(e)}")
            return False, f"Error checking COM3: {str(e)}"

    def force_com3_reset(self):
        """Forcefully reset COM3."""
        try:
            self.status_label.config(text="Status: Forcing COM3 Reset...")
            logging.info("Initiating force reset of COM3")

            is_free, status_msg = self.check_com3_status()
            if not is_free:
                logging.warning(status_msg)
                messagebox.showwarning("Warning", status_msg + ". Attempting to free COM3.")

            telescope_apps = ["cpwi.exe", "stellarium.exe", "skyportal.exe", "serialport.exe"]
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.name().lower() in telescope_apps:
                        proc.kill()
                        logging.info(f"Killed {proc.name()} (PID {proc.pid})")
                        time.sleep(1)
                except (psutil.NoSuchProcess, psutil.Error):
                    continue

            handle_path = r"C:\Sysinternals\handle.exe"
            if os.path.exists(handle_path):
                try:
                    result = subprocess.run(
                        [handle_path, "-a", "COM3"], capture_output=True, text=True
                    )
                    for line in result.stdout.splitlines():
                        if "pid:" in line.lower():
                            try:
                                pid = int(line.split("pid:")[1].split()[0])
                                proc = psutil.Process(pid)
                                proc.kill()
                                logging.info(f"Killed PID {pid} locking COM3")
                                time.sleep(1)
                            except (ValueError, psutil.Error):
                                continue
                except subprocess.SubprocessError as e:
                    logging.error(f"Handle.exe failed: {str(e)}")

            ps_usb_reset = """
            $dev = Get-PnpDevice | Where-Object { $_.FriendlyName -like "*Prolific*USB-to-Serial*" -or $_.FriendlyName -like "*COM3*" -or $_.InstanceId -like "*USB\\VID_067B&PID_2303*" }
            if ($dev) {
                $dev | Disable-PnpDevice -Confirm:$false
                Start-Sleep -Seconds 3
                $dev | Enable-PnpDevice -Confirm:$false
                Write-Output "Reset USB device for COM3"
            } else {
                Write-Output "No matching USB device found for COM3"
            }
            """
            try:
                with open("reset_com3_usb.ps1", "w") as f:
                    f.write(ps_usb_reset)
                result = subprocess.run(
                    ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", "reset_com3_usb.ps1"],
                    capture_output=True, text=True
                )
                os.remove("reset_com3_usb.ps1")
                if result.returncode == 0:
                    logging.info(f"COM3 USB reset: {result.stdout}")
                else:
                    logging.error(f"COM3 USB reset failed: {result.stderr}")
            except Exception as e:
                logging.error(f"PowerShell USB reset failed: {str(e)}")

            try:
                result = subprocess.run(
                    ["pnputil", "/restart-device", "USB\\VID_067B&PID_2303"],
                    capture_output=True, text=True
                )
                if result.returncode == 0:
                    logging.info(f"pnputil restarted Prolific driver: {result.stdout}")
                else:
                    logging.error(f"pnputil restart failed: {result.stderr}")
            except Exception as e:
                logging.error(f"pnputil restart failed: {str(e)}")

            devcon_path = r"C:\Program Files (x86)\Windows Kits\10\Tools\x64\devcon.exe"
            if os.path.exists(devcon_path):
                try:
                    result = subprocess.run(
                        [devcon_path, "restart", "PORT\\COM3"], capture_output=True, text=True
                    )
                    if result.returncode == 0:
                        logging.info("COM3 restarted via devcon")
                    else:
                        logging.error(f"Devcon restart failed: {result.stderr}")
                except Exception as e:
                    logging.error(f"Devcon restart failed: {str(e)}")

            if self.reset_com_port("COM3"):
                logging.info("COM3 reset successful")
                messagebox.showinfo("Success", "COM3 reset completed. Try Test ASCOM or Initialize Mount.")
            else:
                logging.error("COM3 reset failed after all methods")
                messagebox.showerror(
                    "Error", "Failed to reset COM3. Unplug/replug USB, update driver, or try another port."
                )
            self.status_label.config(text="Status: COM3 Reset Complete")
        except Exception as e:
            logging.error(f"Force COM3 reset failed: {str(e)}")
            messagebox.showerror("Error", f"Failed to reset COM3: {str(e)}")

    def reset_com_port(self, port):
        """Reset COM port without mode command."""
        try:
            for attempt in range(3):
                if wmi:
                    try:
                        c = wmi.WMI()
                        for proc in c.Win32_Process():
                            try:
                                cmdline = proc.CommandLine or ""
                                if port.lower() in cmdline.lower() or port.lower() in proc.Name.lower():
                                    proc.Terminate()
                                    logging.info(f"Terminated {proc.Name} (PID {proc.ProcessId}) locking {port}")
                                    time.sleep(1)
                            except Exception:
                                continue
                    except Exception as e:
                        logging.error(f"WMI scan failed: {str(e)}")

                try:
                    for proc in psutil.process_iter(["pid", "name"]):
                        try:
                            for handle in proc.open_files():
                                if port.lower() in handle.path.lower():
                                    proc.terminate()
                                    logging.info(f"Terminated {proc.name()} (PID {proc.pid}) locking {port}")
                                    time.sleep(1)
                                    return True
                        except (psutil.NoSuchProcess, psutil.AccessDenied):
                            continue
                except Exception as e:
                    logging.error(f"psutil scan failed: {str(e)}")

                devcon_path = r"C:\Program Files (x86)\Windows Kits\10\Tools\x64\devcon.exe"
                if os.path.exists(devcon_path):
                    try:
                        result = subprocess.run(
                            [devcon_path, "restart", f"PORT\\{port}"], capture_output=True, text=True
                        )
                        if result.returncode == 0:
                            logging.info(f"Reset {port} using devcon")
                            return True
                        else:
                            logging.warning(f"Devcon reset failed for {port}: {result.stderr}")
                    except Exception as e:
                        logging.error(f"Devcon reset failed for {port}: {str(e)}")

                time.sleep(2 ** attempt)
            return False
        except Exception as e:
            logging.error(f"Failed to reset {port}: {str(e)}")
            return False

    def fix_com_port_permissions(self, port):
        """Fix COM3 permissions with registry and icacls."""
        try:
            try:
                ps_policy_check = subprocess.run(
                    ["powershell.exe", "-Command", "Get-ExecutionPolicy"],
                    capture_output=True, text=True
                )
                if "Restricted" in ps_policy_check.stdout:
                    subprocess.run(
                        ["powershell.exe", "-Command", "Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted -Force"],
                        capture_output=True
                    )
                    logging.info("Set PowerShell policy to Unrestricted")
            except Exception as e:
                logging.warning(f"Error setting PowerShell policy: {str(e)}")

            ps_registry = f"""
            $port = '{port}'
            $paths = @(
                "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\USB\\VID_067B&PID_2303",
                "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\COM Name Arbiter",
                "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\ser2pl\\Parameters",
                "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\Serenum",
                "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\Ports"
            )
            foreach ($path in $paths) {{
                try {{
                    $items = Get-Item -Path $path -ErrorAction SilentlyContinue
                    if ($items) {{
                        foreach ($item in $items) {{
                            $acl = Get-Acl -Path $item.PSPath
                            $rule = New-Object System.Security.AccessControl.RegistryAccessRule("Everyone","FullControl","Allow")
                            $acl.SetAccessRule($rule)
                            Set-Acl -Path $item.PSPath -AclObject $acl
                            Write-Output "Permissions granted to $($item.PSPath)"
                        }}
                    }}
                }} catch {{
                    Write-Error "Failed to set permissions for $path"
                }}
            }}
            """
            try:
                with open("fix_com_permissions.ps1", "w") as f:
                    f.write(ps_registry)
                result = subprocess.run(
                    ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", "fix_com_permissions.ps1"],
                    capture_output=True, text=True
                )
                os.remove("fix_com_permissions.ps1")
                if result.returncode == 0:
                    logging.info(f"Registry permissions granted for {port}: {result.stdout}")
                else:
                    logging.error(f"Registry permissions failed for {port}: {result.stderr}")
            except Exception as e:
                logging.error(f"PowerShell registry permission fix failed: {str(e)}")

            try:
                icacls_cmd = f"icacls \\\\.\\{port} /grant Everyone:F"
                result = subprocess.run(
                    icacls_cmd, shell=True, capture_output=True, text=True
                )
                if result.returncode == 0:
                    logging.info(f"Granted access to {port} via icacls: {result.stdout}")
                else:
                    logging.error(f"icacls failed for {port}: {result.stderr}")
            except Exception as e:
                logging.error(f"icacls permission fix failed: {str(e)}")

            is_free, _ = self.check_com3_status()
            if is_free:
                messagebox.showinfo("Success", f"Permissions granted to {port}. Try again.")
                return True
            else:
                messagebox.showerror("Error", f"Failed to fix {port}. Unplug/replug USB or update driver.")
                return False
        except Exception as e:
            logging.error(f"Error fixing {port} permissions: {str(e)}")
            messagebox.showerror("Error", f"Failed to fix {port}: {str(e)}")
            return False

    def test_ascom(self):
        """Test ASCOM connection with unpark and COM3 retries."""
        if not win32com:
            messagebox.showerror("Error", "win32com.client not available. Install pywin32.")
            return
        try:
            selected_port = self.serial_port_var.get()
            logging.info(f"Testing ASCOM on {selected_port}")

            is_free, status_msg = self.check_com3_status()
            if not is_free:
                logging.error(status_msg)
                messagebox.showerror("Error", status_msg + ". Run 'Force COM3 Reset'.")
                return

            for attempt in range(10):
                try:
                    test_serial = serial.Serial(selected_port, 9600, timeout=0.5)
                    test_serial.close()
                    logging.info(f"{selected_port} access test passed for ASCOM, attempt {attempt + 1}")
                    break
                except serial.SerialException as e:
                    logging.error(f"{selected_port} test failed for ASCOM, attempt {attempt + 1}: {str(e)}")
                    if "access denied" in str(e).lower() or "PermissionError" in str(e).lower():
                        self.reset_com_port(selected_port)
                        if not self.fix_com_port_permissions(selected_port):
                            self.force_com3_reset()
                        time.sleep(2 ** min(attempt, 3))
                    else:
                        raise Exception(f"{selected_port} error: {str(e)}")
            else:
                raise Exception(f"{selected_port} access failed after 10 attempts for ASCOM")

            telescope = win32com.client.Dispatch(self.ascom_driver)
            telescope.Connected = True
            if telescope.Connected:
                try:
                    if hasattr(telescope, "CanUnpark") and telescope.CanUnpark and telescope.Parked:
                        telescope.Unpark()
                        logging.info("Telescope unparked during test")
                    can_move_axis = telescope.CanMoveAxis(0) and telescope.CanMoveAxis(1)
                    logging.info(f"ASCOM {self.ascom_driver} connected, CanMoveAxis={can_move_axis}")
                except Exception as e:
                    logging.error(f"Failed to check CanUnpark or CanMoveAxis: {str(e)}")
                telescope.Connected = False
                messagebox.showinfo("Success", f"ASCOM driver {self.ascom_driver} connected")
            else:
                raise Exception("ASCOM connection failed")
        except Exception as e:
            logging.error(f"ASCOM test failed: {str(e)}")
            messagebox.showerror("Error", f"ASCOM test failed: {str(e)}. Try 'Force COM3 Reset' or unpark via CPWI.")

    def initialize_mount(self):
        """Initialize the telescope mount via serial or ASCOM."""
        try:
            self.status_label.config(text="Status: Initializing Mount...")
            logging.info("Starting mount initialization")
            messagebox.showinfo(
                "Info",
                "Close all telescope apps (e.g., CPWI, SkyPortal). Verify mount on COM3 in Device Manager, then click OK."
            )

            self.cleanup_mount()

            available_ports = []
            for port in serial.tools.list_ports.comports():
                try:
                    port_info = f"{port.device}: {port.description}, VID={port.vid:04X}, PID={port.pid:04X}"
                    available_ports.append(port.device)
                    logging.info(f"Port info: {port_info}")
                except Exception as e:
                    logging.warning(f"Port info error for {port.device}: {str(e)}")
                    available_ports.append(port.device)

            mount_status = "Failed"
            try:
                selected_port = self.serial_port_var.get()
                ports_to_try = [selected_port] + [p for p in available_ports if p != selected_port]
                for port in ports_to_try:
                    if port not in available_ports:
                        logging.warning(f"Skipping invalid port: {port}")
                        continue
                    try:
                        for attempt in range(10):
                            try:
                                self.reset_com_port(port)
                                self.mount = serial.Serial(port, 9600, timeout=0.5)
                                self.send_mount_command(b"V")
                                response = self.mount.read(10)
                                if response:
                                    logging.info(f"Mount connected on {port}")
                                    self.serial_port_var.set(port)
                                    mount_status = f"OK (Serial on {port})"
                                    self.mount_initialized = True
                                    break
                                else:
                                    if self.mount:
                                        self.mount.close()
                                    self.mount = None
                            except serial.SerialException as e:
                                logging.error(f"Serial error on {port}, attempt {attempt + 1}: {str(e)}")
                                if self.mount:
                                    self.mount.close()
                                self.mount = None
                                if "access denied" in str(e).lower() or "PermissionError" in str(e).lower():
                                    self.reset_com_port(port)
                                    if self.fix_com_port_permissions(port):
                                        continue
                                    self.force_com3_reset()
                                time.sleep(2 ** min(attempt, 3))
                        if self.mount:
                            break
                    except Exception as e:
                        logging.error(f"Serial mount failed on {port}: {str(e)}")
            except Exception as e:
                logging.warning(f"Serial mount failed: {str(e)}")

            if not self.mount and win32com:
                try:
                    logging.info("Serial mount failed, trying ASCOM")
                    self.reset_com_port(selected_port)
                    self.ascom_telescope = win32com.client.Dispatch(self.ascom_driver)
                    self.ascom_telescope.Connected = True
                    if self.ascom_telescope.Connected:
                        try:
                            if hasattr(self.ascom_telescope, "CanUnpark") and self.ascom_telescope.CanUnpark:
                                if self.ascom_telescope.Parked:
                                    self.ascom_telescope.Unpark()
                                    logging.info("Mount unparked")
                            self.ascom_telescope.Tracking = False
                            try:
                                can_move_axis = self.ascom_telescope.CanMoveAxis(0) and self.ascom_telescope.CanMoveAxis(1)
                                logging.info(f"ASCOM connected: {self.ascom_driver}, CanMoveAxis={can_move_axis}")
                            except Exception as e:
                                logging.warning(f"Failed to check CanMoveAxis: {str(e)}. Assuming supported")
                            mount_status = "OK (ASCOM)"
                            self.mount_initialized = True
                        except Exception as e:
                            logging.error(f"ASCOM setup failed: {str(e)}")
                            self.ascom_telescope = None
                            raise
                    else:
                        raise Exception("ASCOM connection failed")
                except Exception as e:
                    logging.error(f"ASCOM connection failed: {str(e)}")
                    self.ascom_telescope = None
                    mount_status = "Failed"

            self.update_mount_controls()
            self.status_label.config(text=f"Status: Mount {mount_status}")
            logging.info(f"Mount initialization: {mount_status}")
        except Exception as e:
            logging.error(f"Mount initialization failed: {str(e)}")
            self.mount = None
            self.ascom_telescope = None
            self.mount_initialized = False
            self.update_mount_controls()
            self.status_label.config(text=f"Status: Mount Failed - {str(e)}")
        finally:
            self.root.update()

    def capture_and_save(self):
        """Capture and save camera and spectrograph data."""
        last_time = time.time()
        while self.capture_running and self.running:
            if not self.camera_running or not self.spectrograph_running:
                logging.warning("Capture stopped: Camera or spectrograph not running")
                self.capture_running = False
                self.recording = False
                self.update_record_button_style()
                self.status_label.config(text="Status: Recording Stopped")
                break
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                image_saved = False
                spectrum_saved = False
                # Capture image
                if self.camera and self.camera.isOpened():
                    try:
                        if self.use_camera_exposure.get():
                            self.camera.set(cv2.CAP_PROP_EXPOSURE, self.cam_exposure_ms / 1000.0)
                            time.sleep(self.cam_exposure_ms / 1000.0)  # Wait for exposure
                        for _ in range(10):  # Retry up to 10 frames
                            ret, frame = self.camera.read()
                            if ret:
                                image_path = os.path.join(IMAGE_DIR, f"image_exp{self.cam_exposure_ms:.1f}ms_{timestamp}.jpg")
                                cv2.imwrite(image_path, frame, [int(cv2.IMWRITE_JPEG_QUALITY), 90])
                                image_saved = True
                                logging.info(f"Saved image: {image_path}")
                                break
                            logging.warning("Failed to capture frame")
                            time.sleep(0.1)
                    except Exception as e:
                        logging.error(f"Image capture failed: {str(e)}")

                # Capture spectrum
                if self.spectrograph:
                    try:
                        wavelengths = self.spectrograph.wavelengths()
                        intensities = self.spectrograph.intensities()
                        if self.dark_spectrum is not None and len(intensities) == len(self.dark_spectrum):
                            intensities = intensities - self.dark_spectrum
                        mask = (wavelengths >= 200) & (wavelengths <= 900)
                        spectrum_df = pd.DataFrame({
                            "wavelength_nm": wavelengths[mask],
                            "intensity": intensities[mask]
                        })
                        spectrum_path = os.path.join(
                            SPECTRA_DIR, f"spectrum_{int(self.spec_integration_ms)}ms_{timestamp}.csv"
                        )
                        spectrum_df.to_csv(spectrum_path, index=False)
                        spectrum_saved = True
                        logging.info(f"Saved spectrum: {spectrum_path}")
                    except Exception as e:
                        logging.error(f"Failed to capture spectrum: {str(e)}")

                if image_saved or spectrum_saved:
                    self.sample_count += 1
                    logging.info(f"Sample {self.sample_count}: Image={image_saved}, Spectrum={spectrum_saved}")

                current_time = time.time()
                if current_time - last_time >= 1.0:
                    logging.info(f"Capture rate: {1/(current_time - last_time):.1f} Hz")
                    last_time = current_time
            except Exception as e:
                logging.error(f"Capture and save failed: {str(e)}")
            time.sleep(0.1)

    def move_mount(self, angle_degrees):
        """Move the mount by specified angle (degrees)."""
        try:
            if self.ascom_telescope:
                try:
                    if not self.ascom_telescope.Connected:
                        raise Exception("Telescope not connected")
                    # Move RA axis
                    rate_degrees_sec = self.scan_speed / 3600.0  # Convert arcsec/sec to deg/sec
                    duration = abs(angle_degrees / rate_degrees_sec)  # Seconds to move
                    ra_rate = rate_degrees_sec if angle_degrees >= 0 else -rate_degrees_sec
                    logging.info(f"Moving RA axis to {angle_degrees:.2f}° at {ra_rate:.6f}°/sec for {duration:.2f}s")
                    self.ascom_telescope.MoveAxis(0, ra_rate)
                    time.sleep(duration)
                    self.ascom_telescope.MoveAxis(0, 0)
                    logging.info("RA axis movement stopped")
                except Exception as e:
                    logging.error(f"ASCOM move failed: {str(e)}")
                    raise
            elif self.mount:
                try:
                    # Convert degrees to steps for serial mount
                    steps = int(round(angle_degrees * 3600 / 15.0))  # Approx arcsec to steps
                    command = f"RA{steps:+06d}"
                    self.send_mount_command(command.encode())
                    logging.info(f"Sent serial command: {command}")
                except Exception as e:
                    logging.error(f"Serial mount command failed: {str(e)}")
                    raise
        except Exception as e:
            logging.error(f"Mount movement failed: {str(e)}")

    def track_moon(self):
        """Track the moon and perform scanning."""
        try:
            if not self.mount_initialized:
                logging.error("No mount initialized")
                messagebox.showerror("Error", "No telescope mount initialized. Click 'Initialize Mount'.")
                self.scanning = False
                self.status_label.config(text="Status: Scan Failed")
                return
            self.scanning = True
            self.motor_speed = self.calculate_motor_speed()
            total_angle_arcsec = self.scan_angle * 3600  # Convert to arcsec
            num_steps = int(total_angle_arcsec / self.scan_step)
            logging.info(f"Starting scan: angle={self.scan_angle:.2f}°, steps={num_steps}, step={self.scan_step:.2f} arcsec, speed={self.scan_speed:.2f} arcsec/s")
            self.status_label.config(text="Status: Scan Started")
            self.root.update()
            for i in range(num_steps):
                if not self.scanning:
                    logging.info("Scan stopped by user")
                    break
                try:
                    current_angle_arcsec = i * self.scan_step
                    correction_degrees = self.apply_correction(current_angle_arcsec / 3600.0)
                    adjusted_angle_deg = current_angle_arcsec / 3600.0 + correction_degrees
                    self.move_mount(adjusted_angle_deg)
                    self.status_label.config(
                        text=f"Status: Scanning (step {i+1}/{num_steps}, angle={adjusted_angle_deg:.2f} deg)"
                    )
                    logging.debug(f"Step {i+1}/{num_steps}, angle={adjusted_angle_deg:.2f} deg")
                    time.sleep(self.scan_step / self.scan_speed)  # Time per step
                except Exception as e:
                    logging.error(f"Scan step {i+1} failed: {str(e)}")
            self.scanning = False
            self.stop_mount_manual()
            self.status_label.config(text=f"Status: Scan completed, {self.sample_count} samples")
            logging.info(f"Scan completed: {self.sample_count} samples")
            self.root.update()
        except Exception as e:
            logging.error(f"Scan failed: {str(e)}")
            messagebox.showerror("Error", f"Scan failed: {str(e)}")

    def test_camera(self):
        """Test the selected camera."""
        try:
            camera_index = int(self.camera_index_var.get())
            test_cap = cv2.VideoCapture(camera_index, cv2.CAP_DSHOW)
            if not test_cap.isOpened():
                raise Exception("Camera failed to open")
            ret, frame = test_cap.read()
            test_cap.release()
            if ret:
                logging.info(f"Camera {camera_index} test successful")
                messagebox.showinfo("Success", f"Camera {camera_index} is working")
            else:
                logging.error(f"Camera {camera_index} failed to capture frame")
                messagebox.showerror("Error", f"Camera {camera_index} failed to capture frame")
        except Exception as e:
            logging.error(f"Camera test failed: {str(e)}")
            messagebox.showerror("Error", f"Camera test failed: {str(e)}")

    def toggle_camera(self):
        """Start or stop the camera."""
        if self.camera_running:
            self.camera_running = False
            if self.camera:
                self.camera.release()
                self.camera = None
            self.camera_button.config(text="Start Camera")
            self.webcam_label.config(text="No camera feed")
            self.status_label.config(text="Status: Camera Stopped")
            logging.info("Camera stopped")
        else:
            try:
                camera_index = int(self.camera_index_var.get())
                self.camera = cv2.VideoCapture(camera_index, cv2.CAP_DSHOW)
                if not self.camera.isOpened():
                    raise Exception("Failed to open camera")
                # Set default properties
                self.camera.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                self.camera.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
                if self.use_camera_exposure.get():
                    self.camera.set(cv2.CAP_PROP_EXPOSURE, self.cam_exposure_ms / 1000.0)
                self.camera_running = True
                self.camera_button.config(text="Stop Camera")
                self.status_label.config(text="Status: Camera Started")
                logging.info(f"Camera {camera_index} started")
                # Start camera feed update
                threading.Thread(target=self.update_camera_feed, daemon=True).start()
            except Exception as e:
                logging.error(f"Failed to start camera: {str(e)}")
                messagebox.showerror("Error", f"Failed to start camera: {str(e)}")
                if self.camera:
                    self.camera.release()
                    self.camera = None

    def update_camera_feed(self):
        """Update the camera feed in the GUI."""
        while self.camera_running and self.running:
            try:
                if self.camera and self.camera.isOpened():
                    ret, frame = self.camera.read()
                    if ret:
                        # Convert frame to RGB and resize
                        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                        frame_resized = cv2.resize(frame_rgb, (320, 240))
                        image = Image.fromarray(frame_resized)
                        photo = ImageTk.PhotoImage(image)
                        self.webcam_label.configure(image=photo)
                        self.webcam_label.image = photo
                    else:
                        logging.warning("Failed to grab frame")
                time.sleep(0.033)
            except Exception as e:
                logging.error(f"Camera feed update failed: {str(e)}")
                break
        if self.camera_running:
            self.toggle_camera()

    def retry_camera(self):
        """Retry camera initialization."""
        self.reset_camera()
        self.detect_cameras()
        self.toggle_camera()

    def reset_camera(self):
        """Reset the camera."""
        if self.camera:
            self.camera_running = False
            self.camera.release()
            self.camera = None
            self.camera_button.config(text="Start Camera")
            self.webcam_label.config(text="No camera feed")
            self.status_label.config(text="Status: Camera Reset")
            logging.info("Camera reset")

    def set_spectrograph_integration(self):
        """Set spectrograph integration time."""
        try:
            integration_ms = float(self.spec_integration_entry.get())
            if integration_ms <= 0:
                raise ValueError("Integration time must be positive")
            if self.spectrograph:
                self.spectrograph.integration_time_micros(int(integration_ms * 1000))
                self.spec_integration_ms = integration_ms
                logging.info(f"Set spectrograph integration to {integration_ms:.1f}ms")
                self.status_label.config(text=f"Status: Spectrograph integration set to {integration_ms:.1f}ms")
            else:
                messagebox.showerror("Error", "Spectrograph not initialized. Click 'Start Spectrograph'.")
        except ValueError as e:
            logging.error(f"Invalid integration time: {str(e)}")
            messagebox.showerror("Error", f"Invalid integration time: {str(e)}")
        except Exception as e:
            logging.error(f"Failed to set integration time: {str(e)}")
            messagebox.showerror("Error", f"Failed to set integration time: {str(e)}")

    def toggle_spectrograph(self):
        """Start or stop the spectrograph."""
        if self.spectrograph_running:
            self.spectrograph_running = False
            if self.spectrograph:
                try:
                    self.spectrograph.close()
                except Exception:
                    pass
                self.spectrograph = None
            self.spectrograph_button.config(text="Start Spectrograph")
            self.status_label.config(text="Status: Spectrograph Stopped")
            logging.info("Spectrograph stopped")
        else:
            try:
                devices = sb.list_devices()
                if not devices:
                    raise Exception("No spectrograph found")
                try:
                    self.spectrograph = sb.Spectrometer(devices[0])
                except Exception as e:
                    logging.error(f"Spectrograph initialization error: {str(e)}")
                    raise Exception(f"Failed to initialize spectrograph: {str(e)}")
                self.spectrograph.integration_time_micros(int(self.spec_integration_ms * 1000))
                self.spectrograph_running = True
                self.spectrograph_button.config(text="Stop Spectrograph")
                self.status_label.config(text="Status: Spectrograph Started")
                logging.info("Spectrograph started")
                # Start spectrum update
                threading.Thread(target=self.update_spectrum, daemon=True).start()
            except Exception as e:
                logging.error(f"Failed to start spectrograph: {str(e)}")
                messagebox.showerror("Error", f"Failed to start spectrograph: {str(e)}")
                self.spectrograph = None

    def update_spectrum(self):
        """Update the spectrum plot."""
        while self.spectrograph_running and self.running:
            try:
                if self.spectrograph:
                    wavelengths = self.spectrograph.wavelengths()
                    intensities = self.spectrograph.intensities()
                    if self.dark_spectrum and len(intensities) == len(self.dark_spectrum):
                        intensities = intensities - self.dark_spectrum
                    self.line.set_data(wavelengths, intensities)
                    self.ax.set_ylim(np.min(intensities) * 0.9, np.max(intensities) * 1.1)
                    self.canvas.draw()
                time.sleep(0.2)
            except Exception as e:
                logging.error(f"Spectrum update failed: {str(e)}")
                break
        if self.spectrograph_running:
            self.toggle_spectrograph()

    def capture_dark(self):
        """Capture a dark spectrum."""
        try:
            if not self.spectrograph or not self.spectrograph_running:
                messagebox.showerror("Error", "Spectrograph not running. Please start spectrograph.")
                return False
            messagebox.showinfo("Confirm", "Please cover the spectrograph sensor and click OK to capture dark spectrum.")
            intensities = self.spectrograph.intensities()
            self.dark_spectrum = intensities
            logging.info("Dark spectrum captured")
            self.status_label.config(text="Status: Dark spectrum captured")
            messagebox.showinfo("Success", "Dark spectrum captured successfully")
            return True
        except Exception as e:
            logging.error(f"Failed to capture dark spectrum: {str(e)}")
            messagebox.showerror("Error", f"Failed to capture dark spectrum: {str(e)}")
            return False

    def reinitialize_mount(self):
        """Reinitialize the mount."""
        try:
            self.cleanup_mount()
            self.initialize_mount()
            logging.info("Mount reinitialization attempted")
        except Exception as e:
            logging.error(f"Failed to reinitialize mount: {str(e)}")
            messagebox.showerror("Error", f"Failed to reinitialize mount: {str(e)}")

    def refresh_gui(self):
        """Refresh the GUI."""
        try:
            self.detect_cameras()
            self.update_mount_controls()
            self.status_label.config(text="Status: GUI refreshed")
            logging.info("GUI refreshed")
        except Exception as e:
            logging.error(f"Failed to refresh GUI: {str(e)}")
            messagebox.showerror("Error", f"Failed to refresh GUI: {str(e)}")

    def test_serial(self):
        """Test serial communication with the mount."""
        try:
            port = self.serial_port_var.get()
            with serial.Serial(port, 9600, timeout=0.5) as ser:
                ser.write(b"V")
                response = ser.read(10)
                if response:
                    logging.info(f"Serial test on {port} successful")
                    messagebox.showinfo("Success", f"Serial communication on {port} successful")
                else:
                    raise Exception("No response from serial port")
        except Exception as e:
            logging.error(f"Serial test failed: {str(e)}")
            messagebox.showerror("Error", f"Serial test failed: {str(e)}. Try 'Force COM3 Reset'.")

    def start_scan(self):
        """Start the scanning process."""
        try:
            if not self.camera_running or not self.spectrograph_running:
                messagebox.showerror("Error", "Camera and spectrograph must be running to start scan.")
                return False
            threading.Thread(target=self.track_moon, daemon=True).start()
            return True
        except Exception as e:
            logging.error(f"Failed to start scan: {str(e)}")
            messagebox.showerror("Error", f"Failed to start scan: {str(e)}")
            return False

    def stop_scan(self):
        """Stop the scanning process."""
        try:
            self.scanning = False
            self.status_label.config(text="Status: Scan stopped")
            logging.info("Scan stopped")
        except Exception as e:
            logging.error(f"Error stopping scan: {str(e)}")
            messagebox.showerror("Error", f"Error stopping scan: {str(e)}")

    def toggle_record(self):
        """Start or stop recording."""
        try:
            if self.recording:
                self.recording = False
                self.capture_running = False
                self.record_button.config(text="Start Recording")
                self.update_record_button_style()
                self.status_label.config(text="Status: Recording stopped")
                logging.info("Recording stopped")
            else:
                if not self.camera_running or not self.spectrograph_running:
                    messagebox.showerror("Error", "Camera and spectrograph must be running to start recording")
                    return
                self.recording = True
                self.capture_running = True
                self.record_button.config(text="Stop Recording")
                self.update_record_button_style()
                self.status_label.config(text="Status: Recording started")
                logging.info("Recording started")
                threading.Thread(target=self.capture_and_save, daemon=True).start()
        except Exception as e:
            logging.error(f"Error toggling recording: {str(e)}")
            messagebox.showerror("Error", f"Failed to toggle recording: {str(e)}")

    def move_mount_manual(self, direction):
        """Move the mount manually."""
        try:
            if not self.ascom_telescope:
                messagebox.showerror("Error", "Mount not initialized.")
                return
            speed = float(self.slew_speed_var.get())
            if direction == "North":
                self.ascom_telescope.MoveAxis(1, speed)  # Dec axis
            elif direction == "South":
                self.ascom_telescope.MoveAxis(1, -speed)
            elif direction == "East":
                self.ascom_telescope.MoveAxis(0, -speed)  # RA axis
            elif direction == "West":
                self.ascom_telescope.MoveAxis(0, speed)
            logging.info(f"Manual move: {direction} at speed {speed:.2f}")
        except Exception as e:
            logging.error(f"Failed to move mount manually: {str(e)}")
            messagebox.showerror("Error", f"Failed to move mount: {str(e)}")

    def stop_mount_manual(self):
        """Stop manual mount movement."""
        try:
            if self.ascom_telescope:
                self.ascom_telescope.MoveAxis(0, 0)
                self.ascom_telescope.MoveAxis(1, 0)
                logging.info("Manual mount movement stopped")
        except Exception as e:
            logging.error(f"Failed to stop manual mount: {str(e)}")
            messagebox.showerror("Error", f"Failed to stop mount: {str(e)}")

    def set_slew_speed(self):
        """Set the slew speed for manual movement."""
        try:
            speed = float(self.slew_speed_var.get())
            if not 1 <= speed <= 9:
                raise ValueError("Speed must be between 1 and 9")
            self.ascom_slew_speed = speed
            self.status_label.config(text=f"Status: Slew speed set to {speed}")
            logging.info(f"Slew speed set to {speed}")
        except ValueError as e:
            logging.error(f"Invalid slew speed: {str(e)}")
            messagebox.showerror("Error", f"Invalid slew speed: {str(e)}")
        except Exception as e:
            logging.error(f"Failed to set slew speed: {str(e)}")
            messagebox.showerror("Error", f"Failed to set slew speed: {str(e)}")

    def cleanup_mount(self):
        """Cleanup mount connections."""
        try:
            if self.mount:
                try:
                    self.mount.close()
                except Exception:
                    pass
                self.mount = None
            if self.ascom_telescope:
                try:
                    self.ascom_telescope.Connected = False
                except Exception:
                    pass
                self.ascom_telescope = None
            self.mount_initialized = False
            logging.info("Mount connections cleaned up")
        except Exception as e:
            logging.error(f"Failed to clean up mount: {str(e)}")
            messagebox.showerror("Error", f"Failed to clean up mount: {str(e)}")

    def send_mount_command(self, command):
        """Send command to serial mount."""
        try:
            if self.mount:
                self.mount.write(command)
                logging.debug(f"Sent mount command: {command.decode('utf-8')}")
        except Exception as e:
            logging.error(f"Failed to send mount command: {command.decode('utf-8')} {str(e)}")
            messagebox.showerror("Error", f"Failed to send command: {str(e)}")

    def calculate_motor_speed(self):
        """Calculate motor speed (placeholder)."""
        return 0.0

    def apply_correction(self, angle_degrees):
        """Apply correction to angle (degrees)."""
        return 0.0

    def cleanup(self):
        """Cleanup the application."""
        try:
            self.running = False
            self.recording = False
            self.capture_running = False
            self.spectrograph_running = False
            self.camera_running = False
            if self.camera:
                self.camera.release()
                self.camera = None
            if self.spectrograph:
                try:
                    self.spectrograph.close()
                except Exception:
                    pass
                self.spectrograph = None
            self.cleanup_mount()
            logging.info("Application cleanup completed")
        except Exception as e:
            logging.error(f"Application cleanup failed: {str(e)}")
            messagebox.showerror("Error", f"Application cleanup failed: {str(e)}")

def main():
    """Main application entry point."""
    app = None
    try:
        root = tk.Tk()
        app = MoonScannerGUI(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"Application failed to start: {str(e)}")
        raise
    finally:
        if app:
            try:
                app.cleanup()
                logging.info("Application cleanup completed in main")
            except Exception as e:
                logging.error(f"Cleanup failed in main: {str(e)}")
        logging.info("Application closed")

if __name__ == "__main__":
    main()
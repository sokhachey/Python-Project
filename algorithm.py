import os
import psutil
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import speedtest
import platform
import socket
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import openpyxl
import getpass
import wmi  
import math
import subprocess
from tkinter import messagebox


# Function to write data to Excel
def write_to_excel(sheet_name, data, file_name="system_data.xlsx"):
    if not os.path.exists(file_name):
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)
    else:
        workbook = openpyxl.load_workbook(file_name)

    if sheet_name not in workbook.sheetnames:
        worksheet = workbook.create_sheet(sheet_name)
    else:
        worksheet = workbook[sheet_name]

    for row_index, row in enumerate(data, start=1):
        for col_index, value in enumerate(row, start=1):
            worksheet.cell(row=row_index, column=col_index, value=value)

    workbook.save(file_name)
    messagebox.showinfo("Success", f"Data saved to {file_name} successfully!")


# Helper function to safely get attributes with a default fallback
def get_attribute(obj, attr, default='N/A'):
    try:
        return getattr(obj, attr, default)
    except AttributeError:
        return default

# Helper function to get hardware details with error handling
def get_hardware_info():
    hardware_info = {}
    w = wmi.WMI()

    try:
        system_info = w.Win32_ComputerSystem()[0]
        hardware_info['serial_number'] = get_attribute(system_info, 'SerialNumber')
        hardware_info['system_model'] = get_attribute(system_info, 'Model')
        hardware_info['system_manufacturer'] = get_attribute(system_info, 'Manufacturer')
    except Exception as e:
        print(f"Error retrieving system information: {e}")
    
    # Check BIOS for serial number if not found in system info
    if hardware_info.get('serial_number') == 'N/A':
        try:
            bios_info = w.Win32_BIOS()[0]
            hardware_info['serial_number'] = get_attribute(bios_info, 'SerialNumber')
        except Exception as e:
            print(f"Error retrieving BIOS information: {e}")
    
    return hardware_info

# Helper function to get system resource details
def get_system_resources():
    resources = {}

    # Get CPU info
    cpu_info = psutil.cpu_percent(interval=1, percpu=True)
    resources['cpu_total_cores'] = psutil.cpu_count(logical=False)  # Total physical cores
    
    # Round up the number of cores in use and free cores (this is typically an integer but rounded for your request)
    resources['cpu_cores_in_use'] = math.ceil(sum(1 for usage in cpu_info if usage > 0))  # Cores in use
    resources['cpu_free_cores'] = math.ceil(resources['cpu_total_cores'] - resources['cpu_cores_in_use'])  # Free cores

    # Get memory info and round up the values
    memory_info = psutil.virtual_memory()
    resources['memory_total'] = math.ceil(memory_info.total / (1024 ** 3))  # Convert to GB and round up
    resources['memory_used'] = math.ceil(memory_info.used / (1024 ** 3))  # Convert to GB and round up
    resources['memory_free'] = math.ceil(memory_info.free / (1024 ** 3))  # Convert to GB and round up

    # Get disk info and round up the values
    disk_info = psutil.disk_usage('/')
    resources['disk_total'] = math.ceil(disk_info.total / (1024 ** 3))  # Convert to GB and round up
    resources['disk_used'] = math.ceil(disk_info.used / (1024 ** 3))  # Convert to GB and round up
    resources['disk_free'] = math.ceil(disk_info.free / (1024 ** 3))  # Convert to GB and round up

    return resources

# Main function to get system info
def get_system_info():
    # Get basic system info
    pc_name = socket.gethostname()
    ip_address = socket.gethostbyname(pc_name)
    os_info = f"{platform.system()} {platform.release()}"
    username = getpass.getuser()

    # Get hardware details (using helper function)
    hardware_info = get_hardware_info()

    # Get system resource details (using helper function)
    resource_info = get_system_resources()

    # Construct the system info list
    system_info = [
        ["PC Name", pc_name],
        ["Username", username],
        ["IP Address", ip_address],
        ["Operating System", os_info],
        ["Serial Number", hardware_info.get('serial_number', 'N/A')],
        ["System Model", hardware_info.get('system_model', 'N/A')],
        ["System Manufacturer", hardware_info.get('system_manufacturer', 'N/A')],
        ["CPU Total Cores", resource_info['cpu_total_cores']],
        ["CPU Cores In Use", resource_info['cpu_cores_in_use']],
        ["CPU Free Cores", resource_info['cpu_free_cores']],
        ["Memory (Total)", f"{resource_info['memory_total']} GB"],
        ["Memory (Used)", f"{resource_info['memory_used']} GB"],
        ["Memory (Free)", f"{resource_info['memory_free']} GB"],
        ["Disk (Total)", f"{resource_info['disk_total']} GB"],
        ["Disk (Used)", f"{resource_info['disk_used']} GB"],
        ["Disk (Free)", f"{resource_info['disk_free']} GB"],
    ]

    return system_info

# Function to Test Wi-Fi Speed
def test_wifi_speed():
    wifi_speed = None
    try:
        st = speedtest.Speedtest()
        st.get_best_server()
        download_speed = st.download() / 1_000_000  # Convert to Mbps
        upload_speed = st.upload() / 1_000_000  # Convert to Mbps
        ping = st.results.ping
        wifi_speed = f"Download: {download_speed:.2f} Mbps\nUpload: {upload_speed:.2f} Mbps\nPing: {ping} ms"
    except Exception as e:
        wifi_speed = f"Error testing speed: {e}"

    return wifi_speed


# Display Wi-Fi Speed Test Results
def display_wifi_speed():
    speed = test_wifi_speed()
    wifi_file_text.delete(1.0, tk.END)  # Clear the text box
    wifi_file_text.insert(tk.END, speed)


# Retrieve Wi-Fi Passwords (Windows Only)
def get_wifi_passwords():
    wifi_passwords = []
    
    try:
        profiles = subprocess.check_output("netsh wlan show profiles").decode("utf-8", errors="backslashreplace").split('\n')
        profile_names = [i.split(":")[1][1:-1] for i in profiles if "All User Profile" in i]

        for profile in profile_names:
            try:
                password_info = subprocess.check_output(f'netsh wlan show profile "{profile}" key=clear', shell=True).decode("utf-8", errors="backslashreplace")
                password = [i.split(":")[1][1:-1] for i in password_info.split("\n") if "Key Content" in i]
                if password:
                    wifi_passwords.append((profile, password[0]))
            except subprocess.CalledProcessError:
                wifi_passwords.append((profile, "No password set"))
    
        if not wifi_passwords:
            return "No Wi-Fi profiles found."
        return wifi_passwords
    except Exception as e:
        return f"Error retrieving Wi-Fi passwords: {e}"

# Function for remove files
def remove_temp_files():
    temp_path = os.getenv('TEMP')
    if not temp_path:
        messagebox.showerror("Error", "TEMP environment variable not found!")
        return
    
    try:
        for root, dirs, files in os.walk(temp_path):
            for file in files:
                try:
                    os.remove(os.path.join(root, file))
                except PermissionError:
                    messagebox.showwarning("Warning", f"Permission denied for {file}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to delete {file}: {str(e)}")
        messagebox.showinfo("Success", "Temporary files removed successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# Shutdown and Restart
def shutdown():
    os.system("shutdown /s /t 1")


def restart():
    os.system("shutdown /r /t 1")


# Update Real-Time Graphs
def update_graphs():
    global cpu_data, memory_data, disk_data
    cpu_data.append(psutil.cpu_percent())
    memory_data.append(psutil.virtual_memory().percent)
    disk_data.append(psutil.disk_usage('/').percent)
    
    if len(cpu_data) > 50:
        cpu_data.pop(0)
        memory_data.pop(0)
        disk_data.pop(0)

    cpu_line.set_ydata(cpu_data)
    mem_line.set_ydata(memory_data)
    disk_line.set_ydata(disk_data)

    canvas.draw()
    root.after(1000, update_graphs)


# Alert for High CPU Usage
def check_alerts():
    cpu_usage = psutil.cpu_percent()
    memory_usage = psutil.virtual_memory().percent
    disk_usage = psutil.disk_usage('/').percent
    
    if cpu_usage > 80:
        messagebox.showwarning("Alert", f"High CPU Usage: {cpu_usage}%")
    if memory_usage > 85:
        messagebox.showwarning("Alert", f"High Memory Usage: {memory_usage}%")
    if disk_usage > 90:
        messagebox.showwarning("Alert", f"High Disk Usage: {disk_usage}%")
    
    root.after(5000, check_alerts)


# GUI Setup
root = tk.Tk()
root.title("Enhanced System Monitoring Tool")
root.geometry("900x650")
root.configure(bg="#f0f0f0")  # Light gray background

# Set up style for widgets
style = ttk.Style()
style.configure("TButton", font=("Helvetica", 12), padding=6)
style.configure("TLabel", font=("Helvetica", 14, "bold"))
style.configure("TNotebook.Tab", font=("Helvetica", 12))

notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Tab 1: Real-Time Monitoring
monitor_tab = ttk.Frame(notebook)
notebook.add(monitor_tab, text="Real-Time Monitoring")

fig, ax = plt.subplots(figsize=(5, 3))
cpu_data, memory_data, disk_data = [0] * 50, [0] * 50, [0] * 50

cpu_line, = ax.plot(cpu_data, label="CPU Usage (%)")
mem_line, = ax.plot(memory_data, label="Memory Usage (%)")
disk_line, = ax.plot(disk_data, label="Disk Usage (%)")

ax.set_ylim(0, 100)
ax.set_title("Real-Time System Resource Monitoring", fontsize=16, fontweight='bold')
ax.legend(loc="upper right")
canvas = FigureCanvasTkAgg(fig, master=monitor_tab)
canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

# Tab 2: System Info
info_tab = ttk.Frame(notebook)
notebook.add(info_tab, text="System Info")

info_text = ScrolledText(info_tab, wrap=tk.WORD, height=15, font=("Helvetica", 12), bg="#f7f7f7")
info_text.pack(fill=tk.BOTH, padx=10, pady=10, expand=True)

ttk.Button(info_tab, text="Refresh Info", command=lambda: info_text.insert(
    tk.END, "\n".join([f"{k}: {v}" for k, v in get_system_info()]) + "\n")).pack(pady=10, fill=tk.X)

ttk.Button(info_tab, text="Export to Excel", command=lambda: write_to_excel("System Info", get_system_info())).pack(pady=5, fill=tk.X)


# Tab 3: Wi-Fi & File Management
wifi_file_tab = ttk.Frame(notebook)
notebook.add(wifi_file_tab, text="Wi-Fi & File Management")

wifi_file_text = ScrolledText(wifi_file_tab, wrap=tk.WORD, height=15, font=("Helvetica", 12), bg="#f7f7f7")
wifi_file_text.pack(fill=tk.BOTH, padx=10, pady=10, expand=True)

# Add buttons for retrieving Wi-Fi passwords, checking Wi-Fi speed, and removing temporary files
ttk.Button(wifi_file_tab, text="Wi-Fi Speed Test", command=lambda: threading.Thread(target=display_wifi_speed).start()).pack(pady=10, fill=tk.X)
ttk.Button(wifi_file_tab, text="Retrieve Wi-Fi Passwords", command=lambda: threading.Thread(target=display_wifi_passwords).start()).pack(pady=10, fill=tk.X)
ttk.Button(wifi_file_tab, text="Remove Temp Files", command=lambda: threading.Thread(target=handle_remove_temp_files).start()).pack(pady=10, fill=tk.X)

# Display Wi-Fi Passwords function
def display_wifi_passwords():
    wifi_passwords = get_wifi_passwords()
    wifi_file_text.delete(1.0, tk.END)  # Clear the text box
    if isinstance(wifi_passwords, list):
        for profile, password in wifi_passwords:
            wifi_file_text.insert(tk.END, f"Profile: {profile}\nPassword: {password}\n\n")
    else:
        wifi_file_text.insert(tk.END, wifi_passwords)  # Display the error message if any

# Handle removal of temporary files
def handle_remove_temp_files():
    message = remove_temp_files()
    wifi_file_text.delete(1.0, tk.END)  # Clear the text box

# Tab 4: Actions
actions_tab = ttk.Frame(notebook)
notebook.add(actions_tab, text="Actions")

ttk.Button(actions_tab, text="Shutdown", command=shutdown).pack(pady=10, fill=tk.X)
ttk.Button(actions_tab, text="Restart", command=restart).pack(pady=10, fill=tk.X)


# Initialize Real-Time Monitoring
update_graphs()
check_alerts()

root.mainloop()

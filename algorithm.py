import os
import psutil
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import speedtest
import socket
import platform
import GPUtil
import openpyxl 


# Function to write data to Excel
def write_to_excel(sheet_name, data, file_name="system_data.xlsx"):
    # Create a workbook and add a sheet
    if not os.path.exists(file_name):
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)
    else:
        workbook = openpyxl.load_workbook(file_name)

    if sheet_name not in workbook.sheetnames:
        worksheet = workbook.create_sheet(sheet_name)
    else:
        worksheet = workbook[sheet_name]

    # Write the data to the sheet
    for row_index, row in enumerate(data, start=1):
        for col_index, value in enumerate(row, start=1):
            worksheet.cell(row=row_index, column=col_index, value=value)

    # Save the workbook
    workbook.save(file_name)
    messagebox.showinfo("Success", f"Data saved to {file_name} successfully!")


# Functions for shutdown
def shutdown():
    os.system("shutdown /s /t 1")


# Function for restart
def restart():
    os.system("shutdown /r /t 1")


# Function get system details
def get_system_details(command):
    """Helper function to run WMIC commands and fetch details."""
    try:
        result = subprocess.check_output(command, shell=True, universal_newlines=True)
        return result.strip().split("\n")[-1]
    except Exception as e:
        return f"Error: {str(e)}"


# Function to get basic system information
def get_basic_info():
    pc_name = socket.gethostname()
    ip_address = socket.gethostbyname(pc_name)
    os_info = f"{platform.system()} {platform.release()}"
    processor_info = platform.processor()
    username = os.getlogin()  # Logged-in username
    return pc_name, ip_address, os_info, processor_info, username


# Function to get additional system details using WMIC
def get_additional_details():
    serial_number = get_system_details("wmic bios get serialnumber")
    system_model = get_system_details("wmic computersystem get model")
    system_manufacturer = get_system_details("wmic computersystem get manufacturer")
    return serial_number, system_model, system_manufacturer


# Function to format system information for display
def format_system_info(pc_name, username, ip_address, os_info, processor_info, serial_number, system_model, system_manufacturer):
    return (
        f"PC Name: {pc_name}\n"
        f"Username: {username}\n"
        f"IP Address: {ip_address}\n"
        f"Operating System: {os_info}\n"
        f"Processor: {processor_info}\n"
        f"Serial Number: {serial_number}\n"
        f"System Model: {system_model}\n"
        f"System Manufacturer: {system_manufacturer}\n"
    )


# Function to prepare system information for Excel
def prepare_system_info_data(pc_name, username, ip_address, os_info, processor_info, serial_number, system_model, system_manufacturer):
    return [
        ["   System Information   "],
        ["PC Name", pc_name],
        ["Username", username],
        ["IP Address", ip_address],
        ["Operating System", os_info],
        ["Processor", processor_info],
        ["Serial Number", serial_number],
        ["System Model", system_model],
        ["System Manufacturer", system_manufacturer]
    ]


# Main function to collect and save system information
def show_system_info():
    # Get basic and additional details
    pc_name, ip_address, os_info, processor_info, username = get_basic_info()
    serial_number, system_model, system_manufacturer = get_additional_details()

    # Format information for display
    info = format_system_info(pc_name, username, ip_address, os_info, processor_info, serial_number, system_model, system_manufacturer)
    
    # Display in the Text widget
    display_in_text_widget(info)

    # Prepare and save data for Excel
    system_info_data = prepare_system_info_data(pc_name, username, ip_address, os_info, processor_info, serial_number, system_model, system_manufacturer)
    write_to_excel("System Info", system_info_data)


# Function to display information in a Text widget
def display_in_text_widget(info):
    system_info_text.delete(1.0, tk.END)
    system_info_text.insert(tk.END, info) 

    
# Function to get memory information in GB
def get_memory_info():
    memory_info = psutil.virtual_memory()
    total_memory = round(memory_info.total / (1024 ** 3), 2)  
    used_memory = round(memory_info.used / (1024 ** 3), 2) 
    free_memory = round(memory_info.available / (1024 ** 3), 2)  
    return total_memory, used_memory, free_memory


# Function to get disk information in GB
def get_disk_info():
    disk_usage = psutil.disk_usage('/')
    total_disk = round(disk_usage.total / (1024 ** 3), 2)  
    used_disk = round(disk_usage.used / (1024 ** 3), 2)  
    free_disk = round(disk_usage.free / (1024 ** 3), 2)  
    return total_disk, used_disk, free_disk


# Function to get CPU information
def get_cpu_info():
    cpu_total = psutil.cpu_count(logical=True)  # Total number of logical CPUs
    cpu_usage_count = round((cpu_total * psutil.cpu_percent(interval=1)) / 100, 2)  # Cores in use
    cpu_free_count = round(cpu_total - cpu_usage_count, 2)  # Available cores
    return cpu_total, cpu_usage_count, cpu_free_count


# Function to get GPU information with memory in GB
def get_gpu_info():
    gpus = GPUtil.getGPUs()
    gpu_info = []
    if gpus:
        for gpu in gpus:
            gpu_info.append({
                "name": gpu.name,
                "total_memory": round(gpu.memoryTotal / 1024, 2), 
                "used_memory": round(gpu.memoryUsed / 1024, 2),   
                "free_memory": round(gpu.memoryFree / 1024, 2),   
                "load": gpu.load * 100  # GPU load percentage
            })
    return gpu_info


# Function to prepare GPU info text
def format_gpu_info(gpu_info):
    if gpu_info:
        gpu_info_text = ""
        for gpu in gpu_info:
            gpu_info_text += (
                f"GPU: {gpu['name']}\n"
                f"  Total Memory: {gpu['total_memory']} GB\n"
                f"  Used Memory: {gpu['used_memory']} GB\n"
                f"  Free Memory: {gpu['free_memory']} GB\n"
                f"  GPU Load: {gpu['load']:.2f}%\n"
            )
    else:
        gpu_info_text = "No GPU detected.\n"
    return gpu_info_text


# Function to prepare resource data for Excel
def prepare_excel_data(cpu_total, cpu_usage_count, cpu_free_count, total_memory, used_memory, free_memory, total_disk, used_disk, free_disk, gpu_info):
    resource_data = [
        ["   Resource Information   "],
        ["CPU Total Cores", cpu_total],
        ["CPU Cores In Use", cpu_usage_count],
        ["CPU Free Cores", cpu_free_count],
        ["Memory (Total)", f"{total_memory} GB"],
        ["Memory (Used)", f"{used_memory} GB"],
        ["Memory (Free)", f"{free_memory} GB"],
        ["Disk (Total)", f"{total_disk} GB"],
        ["Disk (Used)", f"{used_disk} GB"],
        ["Disk (Free)", f"{free_disk} GB"],
    ]
    if gpu_info:
        for gpu in gpu_info:
            resource_data.append([f"GPU {gpu['name']} Load", f"{gpu['load']:.2f}%"])
            resource_data.append([f"GPU {gpu['name']} Total Memory", f"{gpu['total_memory']} GB"])
            resource_data.append([f"GPU {gpu['name']} Used Memory", f"{gpu['used_memory']} GB"])
            resource_data.append([f"GPU {gpu['name']} Free Memory", f"{gpu['free_memory']} GB"])
    return resource_data


# Main function to combine all resources
def check_resources():
    total_memory, used_memory, free_memory = get_memory_info()
    total_disk, used_disk, free_disk = get_disk_info()
    cpu_total, cpu_usage_count, cpu_free_count = get_cpu_info()
    gpu_info = get_gpu_info()

    # Format GPU info text
    gpu_info_text = format_gpu_info(gpu_info)

    # Combine all information for display
    resource_info = (
        f"CPU:\n"
        f"  Total Cores: {cpu_total}\n"
        f"  In Use: {cpu_usage_count} Cores\n"
        f"  Free: {cpu_free_count} Cores\n\n"
        f"Memory (RAM):\n"
        f"  Total: {total_memory} GB\n"
        f"  Used: {used_memory} GB\n"
        f"  Free: {free_memory} GB\n\n"
        f"Disk (Hard Disk):\n"
        f"  Total: {total_disk} GB\n"
        f"  Used: {used_disk} GB\n"
        f"  Free: {free_disk} GB\n\n"
        f"{gpu_info_text}"
    )

    # Display in Text widget
    display_in_text_widget(resource_info)

    # Prepare and save data to Excel
    resource_data = prepare_excel_data(cpu_total, cpu_usage_count, cpu_free_count, total_memory, used_memory, free_memory, total_disk, used_disk, free_disk, gpu_info)
    write_to_excel("Resource Info", resource_data)


# Function to display resource info in a Text widget
def display_in_text_widget(resource_info):
    system_info_text.delete(1.0, tk.END)
    system_info_text.insert(tk.END, resource_info)
import os
import psutil
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import speedtest        # For checking Wi-Fi speed.
import socket
import platform
                                                                                                                                                                                                                                                                                                                                              
# Functions for the application
def shutdown():
    os.system("shutdown /s /t 1")

def restart():
    os.system("shutdown /r /t 1")

def show_system_info():
    pc_name = socket.gethostname()
    ip_address = socket.gethostbyname(pc_name)
    os_info = f"{platform.system()} {platform.release()}"
    processor_info = platform.processor()

    info = (
        f"PC Name: {pc_name}\n"
        f"IP Address: {ip_address}\n"
        f"Operating System: {os_info}\n"
        f"Processor: {processor_info}\n"
    )
    
    system_info_text.delete(1.0, tk.END)
    system_info_text.insert(tk.END, info)

# GUI Setup
root = tk.Tk()
root.title("System Management Tool")
root.geometry("600x500")

# Buttons
ttk.Button(root, text="Shutdown", command=shutdown).pack(pady=5)
ttk.Button(root, text="Restart", command=restart).pack(pady=5)
ttk.Button(root, text="Show System Info", command=show_system_info).pack(pady=5)
# Text Box to Display Output
system_info_text = ScrolledText(root, wrap=tk.WORD, width=70, height=15)
system_info_text.pack(pady=10)

# Run the application
root.mainloop()








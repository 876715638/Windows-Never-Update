import ctypes
import sys
import win32serviceutil
import win32service
import tkinter as tk
from tkinter import Tk, Button, messagebox
import winreg as reg
import subprocess


# Refresh wuauserv service registry path
def refresh_wuauserv_registry_path():
    original_service_name = "wuauserv"
    original_service_path = get_service_path(original_service_name)
    if original_service_path:
        print(f"Service '{original_service_name}' binary path: {original_service_path}")
    else:
        print(f"Failed to get the path for service '{original_service_name}'")
    service_path = r"SYSTEM\CurrentControlSet\Services\wuauserv"
    value_name = "ImagePath"
    new_path = r"C:\Windows\system32\svchost.exe -k netsvcs -p"
    try:
        # Open the registry key with write access
        key = reg.OpenKey(reg.HKEY_LOCAL_MACHINE, service_path, 0, reg.KEY_WRITE)

        # Set the new executable path
        reg.SetValueEx(key, value_name, 0, reg.REG_SZ, new_path)

        # Close the registry key
        reg.CloseKey(key)

        print(f"Path of 'wuauserv' service has been changed to: {new_path}")
    except PermissionError:
        print("Permission denied. Try running the script as an administrator.")
    except Exception as e:
        print(f"Error: {e}")

# Admin check functions
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False

def restart_as_admin():
    """如果没有管理员权限，则以管理员身份重新启动程序"""
    if not is_admin():
        # Restart the program by runas and run as admin
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
        sys.exit(0)

# Service control functions
def stop_service():
    service_name = "wuauserv"
    try:
        win32serviceutil.StopService(service_name)
        messagebox.showinfo("Success", f"{service_name} has been stopped.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to stop {service_name}: {e}")

def start_service():
    refresh_wuauserv_registry_path()
    service_name = "wuauserv"
    try:
        win32serviceutil.StartService(service_name)
        messagebox.showinfo("Success", f"{service_name} has been started.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start {service_name}: {e}")

def disable_service():
    service_name = "wuauserv"
    try:
        win32serviceutil.ChangeServiceConfig(
            pythonClassString=None,
            serviceName=service_name,
            startType=win32service.SERVICE_DISABLED,
            errorControl=win32service.SERVICE_ERROR_IGNORE,
            password=None,
            displayName=None
        )
        messagebox.showinfo("Success", f"{service_name} has been disabled.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to disable {service_name}: {e}")

def enable_service_auto():
    refresh_wuauserv_registry_path()
    service_name = "wuauserv"
    try:
        win32serviceutil.ChangeServiceConfig(
            pythonClassString=None,
            serviceName=service_name,
            startType=win32service.SERVICE_AUTO_START,
            errorControl=win32service.SERVICE_ERROR_IGNORE,
            password=None,
            displayName=None,
        )
        messagebox.showinfo("Success", f"{service_name} has been enabled.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to enable {service_name}: {e}")

def enable_service_manual():
    refresh_wuauserv_registry_path()
    service_name = "wuauserv"
    try:
        win32serviceutil.ChangeServiceConfig(
            pythonClassString=None,
            serviceName=service_name,
            startType=win32service.SERVICE_DEMAND_START,
            errorControl=win32service.SERVICE_ERROR_IGNORE,
            password=None,
            displayName=None,
        )
        messagebox.showinfo("Success", f"{service_name} has been enabled.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to enable {service_name}: {e}")

#Pull down menu function
def on_selection_change(selected_value):
    print(f"Selected value: {selected_value}")

def get_service_path(service_name):
    try:
        # Use sc qc command to query the service configuration
        result = subprocess.run(["sc", "qc", service_name], capture_output=True, text=True, check=True)
        
        # Check the path line and extract the path
        for line in result.stdout.splitlines():
            if "BINARY_PATH_NAME" in line:
                path = line.split("    ", 1)[1]  # 分割并获取路径部分
                return path
        return None
    except subprocess.CalledProcessError as e:
        print(f"Error occurred: {e}")
        return None

# Create GUI
def create_gui():
    # Get the language pack for each button based on the selected option
    def update_button_text(*args):
        selected_value_start = btn_start_language_pack[language_options.index(selected_option.get())]
        selected_value_stop = btn_stop_language_pack[language_options.index(selected_option.get())]
        selected_value_enable_auto = btn_enable_auto_language_pack[language_options.index(selected_option.get())]
        selected_value_enable_manual = btn_enable_manual_language_pack[language_options.index(selected_option.get())]
        selected_value_disable = btn_disable_language_pack[language_options.index(selected_option.get())]
        btn_start.config(text=f"{selected_value_start}")
        btn_stop.config(text=f"{selected_value_stop}")
        btn_enable_auto.config(text=f"{selected_value_enable_auto}")
        btn_enable_manual.config(text=f"{selected_value_enable_manual}")
        btn_disable.config(text=f"{selected_value_disable}")
    
    root = Tk()
    root.title("Windows Upgrade Service Control")
    root.geometry("325x325")

    #Language Selection
    language_options = ["English", "中文"]

    selected_option = tk.StringVar()
    selected_option.set(language_options[0]) # Default to English

    btn_start_language_pack = ["Start Service (Please Enable Service first)", "开始服务（请先启动服务）"]
    btn_stop_language_pack = ["Stop Service", "停止服务"]
    btn_enable_auto_language_pack = ["Enable Service (Auto)", "启动服务（自动）"]
    btn_enable_manual_language_pack = ["Enable Service (Manual)", "启动服务（手动）"]
    btn_disable_language_pack = ["Disable Service", "禁用服务"]

    # Create OptionMenu Pulldown menu
    dropdown = tk.OptionMenu(root, selected_option, *language_options)
    dropdown.config(bg="lightgreen")
    dropdown.pack(pady=20)
    
    # Add button
    btn_start = Button(root, text=f"{btn_start_language_pack[language_options.index(selected_option.get())]}", width=35, fg="red", command=start_service)
    btn_start.pack(pady=10)

    btn_stop = Button(root, text=f"{btn_stop_language_pack[language_options.index(selected_option.get())]}", width=35, command=stop_service)
    btn_stop.pack(pady=10)

    btn_enable_auto = Button(root, text=f"{btn_enable_auto_language_pack[language_options.index(selected_option.get())]}", width=35, command=enable_service_auto)
    btn_enable_auto.pack(pady=10)

    btn_enable_manual = Button(root, text=f"{btn_enable_manual_language_pack[language_options.index(selected_option.get())]}", width=35, command=enable_service_manual)
    btn_enable_manual.pack(pady=10)

    btn_disable = Button(root, text=f"{btn_disable_language_pack[language_options.index(selected_option.get())]}", width=35, command=disable_service)
    btn_disable.pack(pady=10)

    # Trace to option change to make sure every time the option changes, the text on the button would change
    selected_option.trace("w", update_button_text)  # "w" means write, i.e., when the value changes

    root.mainloop()

# Main function
if __name__ == "__main__":
    # Ask for admin privileges if already running as admin
    if not is_admin():
        restart_as_admin()

    # Create and show GUI
    create_gui()
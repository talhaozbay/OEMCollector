import platform
import psutil
import wmi
import smtplib
import subprocess
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from colorama import init, Fore
import pyfiglet
import tkinter as tk
from tkinter import messagebox
import time
import sys
import threading
import msvcrt


init(autoreset=True)


def loading_spinner(event, message="Processing..."):
    spinner_chars = ['/', '-', '\\', '|']
    i = 0
    while not event.is_set():
        sys.stdout.write(f'\r{message} {spinner_chars[i % len(spinner_chars)]}')
        sys.stdout.flush()
        i += 1
        time.sleep(0.1)
    
    sys.stdout.write('\r' + ' ' * (len(message) + 2 + len(spinner_chars)) + '\r')
    sys.stdout.flush()
    result = "The system OEMKEY has been successfully added to the inventory."
    print(Fore.GREEN + result)

def get_user_info():
    banner = pyfiglet.figlet_format("OEMKEY COLLECTOR")
    print(Fore.GREEN + banner)

    ad = input(Fore.GREEN + "Your Name : ").strip()
    soyad = input(Fore.GREEN + "Your Subname : ").strip()
    return f"{ad} {soyad}"

def get_OEMKEY():
    command = r'''powershell "Get-WmiObject -query 'select * from SoftwareLicensingService' | Select-Object -ExpandProperty OA3xOriginalProductKey"'''
    result = subprocess.check_output(command, shell=True)
    print(Fore.GREEN + "----------------------------------------------------------------------------------------")
    print(result.decode().strip())
    print(Fore.GREEN + "----------------------------------------------------------------------------------------\n")
    return result.decode().strip()

def get_system_info():
    c = wmi.WMI()
    info = "=== System Information ===\n\n"

    for os in c.Win32_OperatingSystem():
        info += f"OS Name: {os.Caption.strip()}\n"

    info += f"System Name: {platform.node()}\n"

    for system in c.Win32_ComputerSystem():
        info += f"System Manufacturer: {system.Manufacturer}\n"
        info += f"System Model: {system.Model}\n"
        info += f"RAM (Total): {round(psutil.virtual_memory().total / (1024 ** 3), 2)} GB\n"

    for processor in c.Win32_Processor():
        info += f"CPU : {processor.Name.strip()}\n"

    for gpu in c.Win32_VideoController():
        info += f"GPU : {gpu.Name}\n"

    info += "\n======== OEM KEY ========\n"

    return info

def send_email(subject, body, to_email, from_email, from_password, done_event):
    message = MIMEMultipart()
    message['From'] = from_email
    message['To'] = to_email
    message['Subject'] = subject

    message.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        server.send_message(message)
        server.quit()
        print("")
    except Exception as e:
        print("The email could not be sent : ", e)
    finally:
        done_event.set()

if __name__ == "__main__":
    
    full_name = get_user_info()

    device_OEM = get_OEMKEY()
    
    done_event = threading.Event()

    spinner_thread = threading.Thread(target=loading_spinner, args=(done_event, Fore.GREEN + "System Information is being sent..."))
    spinner_thread.start()

    system_info = get_system_info()

    mail_body = f"User: {full_name}\n\n{system_info}\n\n{device_OEM}"

    TO_EMAIL = "<your mail>"
    FROM_EMAIL = "<your mail again>"
    FROM_PASSWORD = "<your mail password>"

    

    send_email(
        subject=f"{full_name} - System Information Report",
        body=mail_body,
        to_email=TO_EMAIL,
        from_email=FROM_EMAIL,
        from_password=FROM_PASSWORD,
        done_event=done_event
    )
    root = tk.Tk()
    root.withdraw()
    print("Press any key to close the program...")
    msvcrt.getch()
    sys.exit()


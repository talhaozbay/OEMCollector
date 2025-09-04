
# OEMCollector

**OEMCollector** is a Python tool that collects hardware and license information on Windows-based systems and reports it via e-mail.  
After asking for the user's name, the script retrieves the OEM license key (OA3xOriginalProductKey) using PowerShell, gathers operating system, manufacturer, model, RAM, CPU, and GPU information via WMI, and then sends all the collected data as an email.

The e‑mail is sent through the `smtp.office365.com` server, and the username and password of the email account must be specified in the code.

---

## 🚀 Features

- **Retrieve OEM key:**  
  The `get_OEMKEY()` function runs a PowerShell command to fetch the Windows `OA3xOriginalProductKey` value and prints it to the console.

- **System hardware information:**  
  The `get_system_info()` function collects OS name, computer name, manufacturer/model, RAM capacity, CPU, and GPU details via WMI.

- **Email reporting:**  
  The collected information is combined by the `send_email()` function and sent to the specified recipient via the Office 365 SMTP server. A loading animation shows the progress of the operation.

- **Compiled application:**  
  The `dist/OEMCollector.exe` executable allows the tool to be used without a Python installation. The `OEMCollector.spec` file provides the configuration to build your own `.exe` file with PyInstaller.

---

## 📦 Requirements

This project only works on **Windows systems**.  
If you don't want to use the executable, you can run the script with Python by installing the dependencies:

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --icon=OEMcollector.ico --hidden-import=pyfiglet.fonts --collect-data=pyfiglet OEMCollector.py
```

The `requirements.txt` file includes the following packages:

- `psutil`
- `wmi`
- `pyfiglet`
- `colorama`

⚠️ The `tkinter` module is included in the Python standard library.

---

## ⚙️ Usage

1. **Edit email information**  
   Inside `OEMCollector.py`, replace the constants `TO_EMAIL`, `FROM_EMAIL`, and `FROM_PASSWORD` with your own email address and password.  
   If you want to use an SMTP server other than Office 365, update the server settings in the `send_email()` function.

2. **Run the script**  
   ```bash
   python OEMCollector.py
   ```

3. **Enter your details**  
   When the program starts, it will ask for your first and last name. After entering them, the program will collect the OEM key and system information, then send them to the specified email address.

4. **Close the program**  
   After the report is sent, the program will wait for any key press before closing.

---

## 📂 Using the compiled version

- Run **OEMCollector.exe** in the `dist` folder to use the tool without installing Python.  
- The icon file **OEMcollector.ico** is used.  
- To build your own executable with PyInstaller, run the following command in an environment where PyInstaller is installed:

```bash
pyinstaller OEMCollector.spec
```

This command generates a new **dist** folder with the settings defined in `OEMCollector.spec` (icon file, name, console mode, etc.).

---

## 🔒 Security & Privacy

- Do not share the email credentials in the script with malicious parties.  
- To prevent unauthorized access, **avoid storing your password directly in the code**. Use environment variables or a configuration file instead.  
- The collected OEM key and hardware information are sensitive. Ensure they are only sent to trusted recipients.

---

## 🤝 Contributing

Pull requests and suggestions are welcome.  
To add a new feature or report a bug, please open an **issue**.


---

## 📧 Contact

If you have any questions or feedback, feel free to reach out via the **GitHub profile**.

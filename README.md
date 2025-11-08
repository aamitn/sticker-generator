# 🏷️ Sticker Generator


A lightweight **PyQt6** based desktop utility powered by **python-docx** for generating printable **UPS and Battery Charger sticker documents** in `.docx` format.  

Easily customize **customer details**, **job numbers**, and **product configurations** through an intuitive GUI with built-in options for **auto-opening** or **auto-printing** the generated file.

[![Download StickerGeneratorSetup.exe](https://img.shields.io/badge/Download-StickerGeneratorSetup.exe-blue?logo=download&style=for-the-badge)](github.com/aamitn/sticker-generator/releases/latest/download/StickerGeneratorSetup.exe)

---

## 🚀 Features

- Generate front and back pages with large white headings.  
- Automatically inserts and scales the sticker image.  
- Dynamically adjusts text size to fit a single line.  
- Options to auto-open and auto-print after generation.  
- Supports dark and light mode themes.  
- Optional “Use Default Printer” feature to skip print dialog.  
- Comes with a Windows installer using **Inno Setup**.

---

## 🧱 Project Structure

```
stickering/
│
├── app.py                # Main application script (PyQt6 GUI)
├── sticker.png            # Sticker image used in document
├── icon.ico               # App icon
├── installer/
│   └── iscript.iss        # Inno Setup installer script
└── dist/
    └── app.exe            # Built executable after PyInstaller
```

---

## ⚙️ Installation (Development Environment)

1. **Create and activate a virtual environment:**

   ```bash
   python -m venv .venv
   # or
   python -m venv venv
   ```

2. **Activate it:**

   ```bash
   .\.venv\Scripts\activate.bat
   # or (Linux/Mac)
   source ./venv/bin/activate
   ```

3. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

   (Your requirements file should include `python-docx` and `PyQt6`.)

---

## 🏗️ Build Executable (PyInstaller)

To generate a standalone `.exe` file:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --icon=icon.ico app.py
```

The output executable will be located in the `dist/` folder.

---

## 📦 Create Windows Installer (Inno Setup)

Once you have the executable (`app.exe`), you can compile the installer.

> ⚠️ **Run Command Prompt as Administrator**

```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" iscript.iss
```

This generates a setup file inside the `installer/output` directory.

---

## 🧩 Installer Script Highlights (`installer/iscript.iss`)

- Copies both `app.exe` and `sticker.png` to the install directory.
- Creates Start Menu and Desktop shortcuts.
- Adds custom app icon and post-install “Launch Sticker Generator” option.

---

## 🖨️ Printing Options

- **Auto-open:** Opens generated `.docx` file after creation.  
- **Auto-print:** Prints automatically using system print dialog or default printer.  
- **Default printer mode:** When enabled, printing bypasses the dialog.

---

## 🧑‍💻 Developer Notes

- All text and colors are chosen to remain readable in both dark and light themes.  
- Uses **python-docx** for document creation.  
- Dynamically adjusts font size for product name and serial number to prevent wrapping.

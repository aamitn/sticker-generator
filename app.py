import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QComboBox, QSpinBox,
    QMessageBox, QGroupBox, QFormLayout, QMainWindow, QMenuBar,
    QCheckBox, QProgressDialog
)
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette, QIntValidator, QAction, QTextDocument, QPageSize
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
import subprocess , webbrowser, platform
from pathlib import Path

# ----------------------------------------
# Global Configuration for Input Limits
# ----------------------------------------
KVA_MIN = 0
KVA_MAX = 9999

UPS_SETS_MIN = 1
UPS_SETS_MAX = 20

UPS_PER_SET_MIN = 1
UPS_PER_SET_MAX = 20

CHARGERS_MIN = 1
CHARGERS_MAX = 20

JOB_OP_MAX = 999999

DOCS_DIR = Path.home() / "Documents" / "Sticker Generator"

# ----------------------------------------
# Backend logic
# ----------------------------------------

def fit_text_to_line(run, text, base_font_size=23, max_chars_one_line=40, min_font_size=14):
    """Shrink font size until text fits in one line."""
    text_length = len(text)
    font_size = base_font_size
    while text_length > max_chars_one_line and font_size > min_font_size:
        font_size -= 1
        max_chars_one_line += 3
    run.font.size = Pt(font_size)
    return font_size


def add_page(doc, side, product_label, customer_name, serial_number, sticker_path):
    heading = doc.add_paragraph()
    run = heading.add_run(side)
    run.font.name = "Calibri"
    run.font.size = Pt(48)
    run.font.bold = True
    run.font.underline = True
    run.font.color.rgb = RGBColor(0, 0, 0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("\n")

    # Sticker image
    if os.path.exists(sticker_path):
        doc.add_picture(sticker_path, width=Inches(6.3))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        doc.add_paragraph("[Sticker image missing]").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")

    # Product name
    p_name_text = f"{product_label} ({customer_name})"
    p_name = doc.add_paragraph(p_name_text)
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_name = p_name.runs[0]
    run_name.font.name = "Calibri"
    run_name.font.bold = True
    run_name.font.color.rgb = RGBColor(0, 0, 0)
    fit_text_to_line(run_name, p_name_text, base_font_size=23)

    # Serial number
    p_serial = doc.add_paragraph(serial_number)
    p_serial.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_serial = p_serial.runs[0]
    run_serial.font.name = "Calibri"
    run_serial.font.bold = True
    run_serial.font.color.rgb = RGBColor(0, 0, 0)
    fit_text_to_line(run_serial, serial_number, base_font_size=23)


# ----------------------------------------
# Worker Thread for DOCX Generation
# ----------------------------------------
class DocxWorker(QThread):
    progress = pyqtSignal(int)      # emit progress percentage
    finished = pyqtSignal(str)      # emit final file path
    error = pyqtSignal(str)         # emit error message

    def __init__(self, main_window, **kwargs):
        super().__init__()
        self.main_window = main_window   # ✅ store main window reference
        self.kwargs = kwargs

    def run(self):
        try:
            product_type = self.kwargs.get("product_type").upper().strip()
            customer_name = self.kwargs.get("customer_name").upper().strip()
            sticker_path = self.kwargs.get("sticker_path")
            job_no = self.kwargs.get("job_no")
            op_no = self.kwargs.get("op_no")
            start_index = self.kwargs.get("start_index", 1)

            doc = Document()
            
            # ---------- Use Fiscal Year in serial_number ----------
            if self.main_window.override_fy_cb.isChecked():
                fy_str = self.main_window.fy_dropdown.currentText()
                fy_input_year = 2000 + int(fy_str.split('-')[0])  # convert '25-26' → 2025
                fy = self.main_window.get_financial_year_from_year(fy_input_year)
            else:
                fy = self.main_window.get_current_financial_year()


            # Determine total pages for progress tracking
            total_pages = 0
            if product_type == "UPS":
                num_sets = self.kwargs.get("num_sets", 1)
                ups_per_set = self.kwargs.get("ups_per_set", 1)
                total_pages = num_sets * (ups_per_set + 1) * 2  # +1 for BYPASS, 2 sides
            else:
                num_chargers = self.kwargs.get("num_chargers", 1)
                total_pages = num_chargers * 2

            current_page = 0

            def add_with_progress(*args, **kwargs):
                nonlocal current_page
                add_page(*args, **kwargs)
                current_page += 1
                percent = int((current_page / total_pages) * 100)
                self.progress.emit(percent)

            if product_type == "UPS":
                num_sets = self.kwargs.get("num_sets", 1)
                ups_per_set = self.kwargs.get("ups_per_set", 1)
                kva_rating = self.kwargs.get("kva_rating")
                for set_idx in range(1, num_sets + 1):
                    ups_list = [f"UPS{i + 1}" for i in range(ups_per_set)]
                    if ups_per_set > 1:
                        ups_list.append("BYPASS")

                    for unit in ups_list:
                        product_label = f"{kva_rating}kVA {unit}"
                        serial_number = f"(SL. NO. : LL/{fy}/{job_no}-OP{op_no}/BYP)" if unit == "BYPASS" else f"(SL. NO. : LL/{fy}/{job_no}-OP{op_no}/{unit})"
                        for side in ["FRONT SIDE", "BACK SIDE"]:
                            add_with_progress(doc, side, product_label, customer_name, serial_number, sticker_path)

            else:
                start = 0 if start_index == 0 else 1
                num_chargers = self.kwargs.get("num_chargers", 1)
                voltage = self.kwargs.get("voltage")
                current = self.kwargs.get("current")
                battery_capacity = self.kwargs.get("battery_capacity")
                charger_type = self.kwargs.get("charger_type")
                battery_type = self.kwargs.get("battery_type")
                for i in range(start, num_chargers + start):
                    index_label = "" if i == 0 else str(i)
                    product_label = f"{voltage}V/{current}A {charger_type} for {battery_capacity}Ah {battery_type} battery"
                    serial_number = f"(SL. NO. : LL/{fy}/{job_no}-OP{op_no}/BCH{index_label})"
                    for side in ["FRONT SIDE", "BACK SIDE"]:
                        add_with_progress(doc, side, product_label, customer_name, serial_number, sticker_path)

            filename = f"Sticker_{customer_name}_{job_no}_{op_no}_{product_type}.docx"
            output_path = str(self.main_window.save_output_path(filename))
            doc.save(output_path)
            self.finished.emit(output_path)

        except Exception as e:
            self.error.emit(str(e))


# ----------------------------------------
# GUI Main Window
# ----------------------------------------
class StickerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sticker Generator Tool")
        self.setWindowIcon(QIcon.fromTheme("document-new"))
        self.setFixedWidth(480)
        self.start_index = 1

        self.init_ui()
        self.apply_adaptive_theme()
        self.auto_load_sticker()
        
        # ----------------------------------------
        # Persisted Settings
        # ----------------------------------------
        self.settings = QSettings("Bitmutex", "StickerGenerator")
        self.load_settings()

    def apply_adaptive_theme(self):
        app_palette = QApplication.instance().palette()
        base_color = app_palette.color(QPalette.ColorRole.Window)
        brightness = (base_color.red() * 0.299 + base_color.green() * 0.587 + base_color.blue() * 0.114)
        is_dark = brightness < 128

        palette = QPalette()
        if is_dark:
            palette.setColor(QPalette.ColorRole.Window, QColor("#121212"))
            palette.setColor(QPalette.ColorRole.Base, QColor("#1E1E1E"))
            palette.setColor(QPalette.ColorRole.Text, QColor("#F2F2F2"))
            palette.setColor(QPalette.ColorRole.Button, QColor("#2D89EF"))
            palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        else:
            palette.setColor(QPalette.ColorRole.Window, QColor("#F5F7FB"))
            palette.setColor(QPalette.ColorRole.Base, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.black)
            palette.setColor(QPalette.ColorRole.Button, QColor("#2F80ED"))
            palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        QApplication.instance().setPalette(palette)
        self.setPalette(palette)
        self.setFont(QFont("Segoe UI", 10))

    def auto_load_sticker(self):
        default_img = os.path.join(os.getcwd(), "sticker.png")
        if os.path.exists(default_img):
            self.sticker_path.setText(default_img)

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)

        # Menu Bar
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)

        # Settings
        settings_menu = menu_bar.addMenu("Settings")
        self.start_0_action = QAction("Start BCH numbering from 0", self, checkable=True)
        self.start_0_action.triggered.connect(self.toggle_start_index)
        settings_menu.addAction(self.start_0_action)
        self.use_default_printer_action = QAction("Use Default Printer", self, checkable=True)
        settings_menu.addAction(self.use_default_printer_action)

        # Edit Menu
        edit_menu = menu_bar.addMenu("Edit")
        open_output_action = QAction("Open Output Path", self)
        open_output_action.triggered.connect(self.open_output_path)
        edit_menu.addAction(open_output_action)
        purge_all_action = QAction("Purge All DOCX", self)
        purge_all_action.triggered.connect(self.purge_all_docx)
        edit_menu.addAction(purge_all_action)

        # Help Menu
        about_menu = menu_bar.addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about)
        about_menu.addAction(about_action)
        update_action = QAction("Check for Update", self)
        update_action.triggered.connect(self.open_github_release)
        about_menu.addAction(update_action)
        
        # Fiscal Year Override
        fy_box = QGroupBox("Fiscal Year")
        fy_layout = QHBoxLayout()
        self.override_fy_cb = QCheckBox("Override Fiscal Year")

        self.fy_dropdown = QComboBox()
        # Populate FY dropdown ±20 years from current FY
        current_year = date.today().year
        fy_list = []

        # FY Dropdown Range
        for y in range(current_year - 20, current_year + 21):
            start = y % 100
            end = (y + 1) % 100
            fy_list.append(f"{start:02d}-{end:02d}")

        self.fy_dropdown.addItems(fy_list)

        # Default state
        self.fy_dropdown.setEnabled(self.override_fy_cb.isChecked())

        # 🔹 This line makes the dropdown responsive to the checkbox
        self.override_fy_cb.toggled.connect(self.fy_dropdown.setEnabled)

        # Optional: set default selection to current FY
        current_fy = self.get_current_financial_year()
        index = fy_list.index(current_fy)
        self.fy_dropdown.setCurrentIndex(index)

        fy_layout.addWidget(self.override_fy_cb)
        fy_layout.addWidget(self.fy_dropdown)
        fy_box.setLayout(fy_layout)
        

        # Customer & Job Details
        customer_box = QGroupBox("Customer && Job Details")
        form1 = QFormLayout()
        self.customer_input = QLineEdit()
        self.job_input = QLineEdit()
        self.job_input.setValidator(QIntValidator(0, JOB_OP_MAX))  
        self.op_input = QLineEdit()
        self.product_type = QComboBox()
        self.product_type.addItems(["UPS", "Battery Charger"])
        form1.addRow("Customer Name:", self.customer_input)
        form1.addRow("Job Number:", self.job_input)
        form1.addRow("OP Number:", self.op_input)
        self.op_input.setValidator(QIntValidator(0, JOB_OP_MAX))  
        form1.addRow("Product Type:", self.product_type)
        customer_box.setLayout(form1)

        # Sticker Selection
        sticker_box = QGroupBox("Sticker Image")
        hbox = QHBoxLayout()
        self.sticker_path = QLineEdit()
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_sticker)
        hbox.addWidget(self.sticker_path)
        hbox.addWidget(browse_btn)
        sticker_box.setLayout(hbox)

        # UPS Fields
        ups_box = QGroupBox("UPS Configuration")
        ups_form = QFormLayout()
        self.num_sets = QSpinBox()
        self.num_sets.setRange(UPS_SETS_MIN, UPS_SETS_MAX)
        self.ups_per_set = QSpinBox()
        self.ups_per_set.setRange(UPS_PER_SET_MIN, UPS_PER_SET_MAX)
        self.kva_rating = QLineEdit()
        self.kva_rating.setValidator(QIntValidator(KVA_MIN, KVA_MAX))
        self.kva_rating.setPlaceholderText("e.g. 30 (do not add kVA)")
        ups_form.addRow("Number of Sets:", self.num_sets)
        ups_form.addRow("UPS per Set:", self.ups_per_set)
        ups_form.addRow("Power Rating (kVA):", self.kva_rating)
        ups_box.setLayout(ups_form)

        # Battery Charger Fields
        charger_box = QGroupBox("Battery Charger Configuration")
        ch_form = QFormLayout()
        self.voltage = QLineEdit()
        self.voltage.setValidator(QIntValidator(1, 1000))
        self.current = QLineEdit()
        self.current.setValidator(QIntValidator(1, 500))
        self.battery_capacity = QLineEdit()
        self.battery_capacity.setValidator(QIntValidator(1, 5000))
        self.charger_type = QComboBox()
        self.charger_type.addItems(["FC", "FC & FCB", "FCBC", "DFCBC"])
        self.battery_type = QComboBox()
        self.battery_type.addItems(["VRLA", "NICAD", "Planté", "Tubular", "Li-Ion", "Li-Po"])
        self.num_chargers = QSpinBox()
        self.num_chargers.setRange(CHARGERS_MIN, CHARGERS_MAX)
        ch_form.addRow("Charger Voltage (V):", self.voltage)
        ch_form.addRow("Charger Current (A):", self.current)
        ch_form.addRow("Battery Capacity (Ah):", self.battery_capacity)
        ch_form.addRow("Charger Type:", self.charger_type)
        ch_form.addRow("Battery Type:", self.battery_type)
        ch_form.addRow("Number of Chargers:", self.num_chargers)
        charger_box.setLayout(ch_form)

        # Options
        options_box = QGroupBox("Post Creation Options")
        opt_layout = QVBoxLayout()

        self.auto_open_cb = QCheckBox("Auto-open file after creation")
        self.auto_open_cb.setChecked(False)  

        self.auto_print_cb = QCheckBox("Auto-print file after creation (default printer, A4)")
        self.auto_print_cb.setChecked(True)  

        opt_layout.addWidget(self.auto_open_cb)
        opt_layout.addWidget(self.auto_print_cb)
        options_box.setLayout(opt_layout)


        # Generate Button
        generate_btn = QPushButton("Generate DOCX")
        generate_btn.clicked.connect(self.generate_docx_threaded)
        generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #2f80ed;
                color: white;
                font-weight: bold;
                padding: 8px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1e6cd4;
            }
        """)
        
        main_layout.addWidget(fy_box)
        main_layout.addWidget(customer_box)
        main_layout.addWidget(sticker_box)
        main_layout.addWidget(ups_box)
        main_layout.addWidget(charger_box)
        main_layout.addWidget(options_box)
        main_layout.addWidget(generate_btn)
        central_widget.setLayout(main_layout)
        self.product_type.currentTextChanged.connect(self.update_visibility)
        self.update_visibility()

    # ---------- Handlers ----------
    def toggle_start_index(self, checked):
        self.start_index = 0 if checked else 1

    def show_about(self):
        about_text = (
            "<div style='font-family:Segoe UI; font-size:10pt; color:#333;'>"
            "<h2 style='color:#2F80ED; margin-bottom:4px;'>Sticker Generator Tool</h2>"
            "<p style='margin:2px 0 8px 0;'>Version: <b>0.4</b></p>"
            "<hr style='border:none; border-top:1px solid #ccc; margin:8px 0;'>"
            "<p style='margin:4px 0;'>Developed by <b>Bitmutex Technologies</b></p>"
            "<p style='margin:4px 0;'>Author: <b>Amit Kumar Nandi</b></p>"
            "<p style='margin:6px 0;'>"
            "For updates, documentation, and releases, visit:<br>"
            "<a href='https://bitmutex.com' style='color:#2F80ED; text-decoration:none;'>"
            "https://bitmutex.com</a>"
            "</p>"
            "<hr style='border:none; border-top:1px solid #ccc; margin:8px 0;'>"
            "<p style='font-size:9pt; color:#777;'>© 2025 Bitmutex Technologies. All rights reserved.</p>"
            "</div>"
        )

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("About – Sticker Generator")
        msg_box.setTextFormat(Qt.TextFormat.RichText)
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setText(about_text)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()


    def update_visibility(self):
        is_ups = self.product_type.currentText() == "UPS"
        for box in self.findChildren(QGroupBox):
            if box.title() == "UPS Configuration":
                box.setVisible(is_ups)
            elif box.title() == "Battery Charger Configuration":
                box.setVisible(not is_ups)

    def browse_sticker(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Sticker Image", "", "Image Files (*.png *.jpg *.jpeg)")
        if file_path:
            self.sticker_path.setText(file_path)

    def open_output_path(self):
        """Open the sticker output folder."""
        try:
            if sys.platform.startswith("win"):
                os.startfile(DOCS_DIR)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(DOCS_DIR)])
            else:
                subprocess.run(["xdg-open", str(DOCS_DIR)])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open output folder:\n{e}")
            
    def save_output_path(self, filename):
        DOCS_DIR.mkdir(parents=True, exist_ok=True)
        return DOCS_DIR / filename
        
    def open_github_release(self):
        try:
            webbrowser.open("https://github.com/aamitn/sticker-generator/releases")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open GitHub releases page:\n{e}")

    def purge_all_docx(self):
        """Delete all .docx files from the output folder."""
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete ALL .docx files in:\n{DOCS_DIR}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            deleted_count = 0
            for file in DOCS_DIR.glob("*.docx"):
                try:
                    file.unlink()
                    deleted_count += 1
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Could not delete {file.name}:\n{e}")

            QMessageBox.information(
                self,
                "Deleted",
                f"Total .docx files deleted: {deleted_count}"
                if deleted_count else "No .docx files found."
            )

    # ---------- DOCX Generation with Progress ----------
    def generate_docx_threaded(self):
        try:
            product_type = self.product_type.currentText()
            sticker_path = self.sticker_path.text().strip()
            customer_name = self.customer_input.text().strip()
            job_no = self.job_input.text().strip()
            op_no = self.op_input.text().strip()

            if not all([sticker_path, customer_name, job_no, op_no]):
                raise ValueError("Please fill all required fields and select a sticker image.")

            kwargs = dict(
                product_type=product_type,
                sticker_path=sticker_path,
                customer_name=customer_name,
                job_no=job_no,
                op_no=op_no,
                start_index=self.start_index,
            )

            if product_type == "UPS":
                kva = int(self.kva_rating.text().strip())
                if not (KVA_MIN <= kva <= KVA_MAX):
                    raise ValueError(f"KVA rating must be between {KVA_MIN} and {KVA_MAX}")
                kwargs.update(
                    kva_rating=kva,
                    num_sets=self.num_sets.value(),
                    ups_per_set=self.ups_per_set.value()
                )
            else:
                kwargs.update(
                    voltage=self.voltage.text().strip(),
                    current=self.current.text().strip(),
                    battery_capacity=self.battery_capacity.text().strip(),
                    charger_type=self.charger_type.currentText(),
                    battery_type=self.battery_type.currentText(),
                    num_chargers=self.num_chargers.value()
                )

            self.progress_dialog = QProgressDialog("Generating DOCX...", "Cancel", 0, 100, self)
            self.progress_dialog.setWindowTitle("Please Wait")
            self.progress_dialog.setWindowModality(Qt.WindowModality.ApplicationModal)
            self.progress_dialog.setMinimumDuration(0)
            self.progress_dialog.show()

            self.worker = DocxWorker(self, **kwargs)
            self.worker.progress.connect(self.progress_dialog.setValue)
            self.worker.finished.connect(self.on_generation_finished)
            self.worker.error.connect(lambda e: QMessageBox.critical(self, "❌ Error", e))
            self.worker.start()

        except Exception as e:
            QMessageBox.critical(self, "❌ Error", str(e))

    def on_generation_finished(self, output):
        self.progress_dialog.close()
        QMessageBox.information(self, "✅ Success", f"Document generated successfully:\n{output}")
        if self.auto_open_cb.isChecked():
            os.startfile(output)
        if self.auto_print_cb.isChecked():
            try:
                use_default_printer = self.use_default_printer_action.isChecked()
                current_platform = platform.system().lower()

                if use_default_printer:
                    # Auto print
                    if current_platform.startswith("windows"):
                        os.startfile(output, "print")
                    elif current_platform.startswith(("linux", "darwin")):
                        subprocess.run(["lp", output])
                    else:
                        QMessageBox.information(self, "Info", f"Automatic printing not supported on {current_platform.title()} yet.")
                else:
                    self.print_docx_via_dialog(output)
            except Exception as e:
                QMessageBox.warning(self, "Print Error", f"Could not print automatically:\n{e}")

    # ---------- Helper Functions ----------
    def get_financial_year_from_year(self, year: int) -> str:
        """Convert a given year to FY string 'YY-YY' format (April-March)."""
        start = year % 100
        end = (year + 1) % 100
        return f"{start:02d}-{end:02d}"

    def get_current_financial_year(self) -> str:
        """Calculate current FY based on today's date."""
        today = date.today()
        year = today.year
        month = today.month
        if month >= 4:
            fy_start = year
        else:
            fy_start = year - 1
        return self.get_financial_year_from_year(fy_start)


    def print_docx_via_dialog(self, docx_path):
        """Show print dialog and print DOCX content via system application."""
        try:
            # Initialize printer and dialog - simplified version
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            
            dialog = QPrintDialog(printer, self)
            dialog.setWindowTitle("Select Printer to Print Sticker")

            if dialog.exec():
                # User confirmed - print using system default application
                if sys.platform.startswith("win"):
                    os.startfile(docx_path, "print")
                elif sys.platform == "darwin":
                    subprocess.run(["open", "-a", "Preview", docx_path])
                else:
                    subprocess.run(["xdg-open", docx_path])

        except Exception as e:
            QMessageBox.warning(self, "Print Error", f"Could not print document:\n{e}")
    
    # ---------- Settings Persistence ----------
    def load_settings(self):
        """Restore checkbox states and preferences"""
        self.auto_open_cb.setChecked(self.settings.value("auto_open", True, bool))
        self.auto_print_cb.setChecked(self.settings.value("auto_print", True, bool))
        self.use_default_printer_action.setChecked(self.settings.value("use_default_printer", True, bool))

    def save_settings(self):
        """Save checkbox states and preferences"""
        self.settings.setValue("auto_open", self.auto_open_cb.isChecked())
        self.settings.setValue("auto_print", self.auto_print_cb.isChecked())
        self.settings.setValue("use_default_printer", self.use_default_printer_action.isChecked())
        
    def closeEvent(self, event):
        """Save settings before app closes"""
        self.save_settings()
        event.accept()
        

# ----------------------------------------
# Entry Point
# ----------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StickerApp()
    window.show()
    sys.exit(app.exec())

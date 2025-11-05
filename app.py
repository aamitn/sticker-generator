import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QFileDialog, QComboBox, QSpinBox,
    QMessageBox, QGroupBox, QFormLayout, QMainWindow, QMenuBar, QCheckBox
)
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette, QAction, QTextDocument, QPageSize
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog
from PyQt6.QtCore import Qt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import subprocess
import platform


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
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        missing = doc.add_paragraph("[Sticker image missing]")
        missing.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    #ADD GAP
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
    p_serial_text = serial_number
    p_serial = doc.add_paragraph(p_serial_text)
    p_serial.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_serial = p_serial.runs[0]
    run_serial.font.name = "Calibri"
    run_serial.font.bold = True
    run_serial.font.color.rgb = RGBColor(0, 0, 0)
    fit_text_to_line(run_serial, p_serial_text, base_font_size=23)


def generate_docx(sticker_path, customer_name, job_number, op_number, product_type,
                  kva_rating=None, num_ups=None,
                  voltage=None, current=None, battery_capacity=None,
                  charger_type=None, battery_type=None, num_chargers=None,
                  start_index=1):
    """Generate DOCX for UPS or Battery Charger."""
    doc = Document()
    product_type = product_type.upper().strip()
    customer_name = customer_name.upper().strip()

    if product_type == "UPS":
        ups_list = [f"UPS{i + 1}" for i in range(num_ups)]
        if num_ups > 1:
            ups_list.append("BYPASS")

        for unit in ups_list:
            product_label = f"{kva_rating}kVA {unit}"
            if unit == "BYPASS":
                serial_number = f"(SL. NO. : LL/25-26/{job_number}-OP{op_number}/BYP)"
            else:
                serial_number = f"(SL. NO. : LL/25-26/{job_number}-OP{op_number}/{unit})"

            for side in ["FRONT SIDE", "BACK SIDE"]:
                add_page(doc, side, product_label, customer_name, serial_number, sticker_path)

    elif product_type in ["BATTERY CHARGER", "CHARGER", "BCH"]:
        start = 0 if start_index == 0 else 1
        end = num_chargers + start
        for i in range(start, end):
            index_label = "" if i == 0 else str(i)
            product_label = f"{voltage}V/{current}A {charger_type} for {battery_capacity}Ah {battery_type} battery"
            serial_number = f"(SL. NO. : LL/25-26/{job_number}-OP{op_number}/BCH{index_label})"

            for side in ["FRONT SIDE", "BACK SIDE"]:
                add_page(doc, side, product_label, customer_name, serial_number, sticker_path)

    else:
        raise ValueError("Invalid product type.")

    output_path = f"Sticker_{customer_name}_{job_number}_{op_number}_{product_type}.docx"
    doc.save(output_path)
    return output_path


# ----------------------------------------
# GUI Main Window
# ----------------------------------------

class StickerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sticker Generator – Bitmutex Tools")
        self.setWindowIcon(QIcon.fromTheme("document-new"))
        self.setFixedWidth(480)

        # Default setting
        self.start_index = 1
        self.init_ui()
        self.apply_adaptive_theme()
        self.auto_load_sticker()

    # ---------- Theme Handling ----------
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

    # ---------- Auto Load Sticker ----------
    def auto_load_sticker(self):
        default_img = os.path.join(os.getcwd(), "sticker.png")
        if os.path.exists(default_img):
            self.sticker_path.setText(default_img)

    # ---------- UI ----------
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)

        # Menu bar
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)

        settings_menu = menu_bar.addMenu("Settings")
        self.start_0_action = QAction("Start BCH numbering from 0", self, checkable=True)
        self.start_0_action.triggered.connect(self.toggle_start_index)
        settings_menu.addAction(self.start_0_action)
        
        # Use Default Printer setting
        self.use_default_printer_action = QAction("Use Default Printer", self, checkable=True)
        settings_menu.addAction(self.use_default_printer_action)

        about_menu = menu_bar.addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about)
        about_menu.addAction(about_action)

        # ---- Customer Info ----
        customer_box = QGroupBox("Customer & Job Details")
        form1 = QFormLayout()
        self.customer_input = QLineEdit()
        self.job_input = QLineEdit()
        self.op_input = QLineEdit()
        self.product_type = QComboBox()
        self.product_type.addItems(["UPS", "Battery Charger"])
        form1.addRow("Customer Name:", self.customer_input)
        form1.addRow("Job Number:", self.job_input)
        form1.addRow("OP Number:", self.op_input)
        form1.addRow("Product Type:", self.product_type)
        customer_box.setLayout(form1)

        # ---- Sticker Selection ----
        sticker_box = QGroupBox("Sticker Image")
        hbox = QHBoxLayout()
        self.sticker_path = QLineEdit()
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_sticker)
        hbox.addWidget(self.sticker_path)
        hbox.addWidget(browse_btn)
        sticker_box.setLayout(hbox)

        # ---- UPS Fields ----
        ups_box = QGroupBox("UPS Configuration")
        ups_form = QFormLayout()
        self.ups_number = QSpinBox()
        self.ups_number.setRange(1, 10)
        self.kva_rating = QLineEdit()
        self.kva_rating.setPlaceholderText("e.g. 30 (do not add kVA)")
        ups_form.addRow("Number of UPS:", self.ups_number)
        ups_form.addRow("Power Rating (kVA):", self.kva_rating)
        ups_box.setLayout(ups_form)

        # ---- Battery Charger Fields ----
        charger_box = QGroupBox("Battery Charger Configuration")
        ch_form = QFormLayout()
        self.voltage = QLineEdit()
        self.current = QLineEdit()
        self.battery_capacity = QLineEdit()
        self.charger_type = QComboBox()
        self.charger_type.addItems(["FC", "FCB", "FCBC", "DFCBC"])
        self.battery_type = QComboBox()
        self.battery_type.addItems(["VRLA", "NICAD"])
        self.num_chargers = QSpinBox()
        self.num_chargers.setRange(1, 10)
        ch_form.addRow("Charger Voltage (V):", self.voltage)
        ch_form.addRow("Charger Current (A):", self.current)
        ch_form.addRow("Battery Capacity (Ah):", self.battery_capacity)
        ch_form.addRow("Charger Type:", self.charger_type)
        ch_form.addRow("Battery Type:", self.battery_type)
        ch_form.addRow("Number of Chargers:", self.num_chargers)
        charger_box.setLayout(ch_form)

        # ---- Options ----
        options_box = QGroupBox("Post Creation Options")
        opt_layout = QVBoxLayout()
        self.auto_open_cb = QCheckBox("Auto-open file after creation")
        self.auto_print_cb = QCheckBox("Auto-print file after creation (default printer, A4)")
        opt_layout.addWidget(self.auto_open_cb)
        opt_layout.addWidget(self.auto_print_cb)
        options_box.setLayout(opt_layout)

        # ---- Generate Button ----
        generate_btn = QPushButton("Generate DOCX")
        generate_btn.clicked.connect(self.generate_docx)
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
        QMessageBox.information(
            self,
            "About",
            "Sticker Generator Tool\n"
            "Version: 0.4\n\n"
            "Developed by Bitmutex Technologies\n"
            "Author: Amit Kumar Nandi\n"
            "Website: https://bitmutex.com"
        )

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

    def generate_docx(self):
        try:
            product_type = self.product_type.currentText()
            sticker_path = self.sticker_path.text().strip()
            customer_name = self.customer_input.text().strip()
            job_no = self.job_input.text().strip()
            op_no = self.op_input.text().strip()

            if not all([sticker_path, customer_name, job_no, op_no]):
                raise ValueError("Please fill all required fields and select a sticker image.")

            if product_type == "UPS":
                kva = self.kva_rating.text().strip()
                num = self.ups_number.value()
                output = generate_docx(sticker_path, customer_name, job_no, op_no, product_type,
                                       kva_rating=kva, num_ups=num)
            else:
                output = generate_docx(sticker_path, customer_name, job_no, op_no, product_type,
                                       voltage=self.voltage.text().strip(),
                                       current=self.current.text().strip(),
                                       battery_capacity=self.battery_capacity.text().strip(),
                                       charger_type=self.charger_type.currentText(),
                                       battery_type=self.battery_type.currentText(),
                                       num_chargers=self.num_chargers.value(),
                                       start_index=self.start_index)

            QMessageBox.information(self, "✅ Success", f"Document generated successfully:\n{output}")

            # --- Auto actions ---
            if self.auto_open_cb.isChecked():
                os.startfile(output)


            if self.auto_print_cb.isChecked():
                try:
                    use_default_printer = self.use_default_printer_action.isChecked()

                    if use_default_printer:
                        # Direct print using default printer (no dialog)
                        if sys.platform.startswith("win"):
                            os.startfile(output, "print")
                        elif sys.platform == "darwin":
                            subprocess.run(["lp", output])
                        else:
                            subprocess.run(["lp", output])
                    else:
                        # Show print dialog for manual printer selection
                        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
                        printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
                        dialog = QPrintDialog(printer, self)
                        dialog.setWindowTitle("Select Printer to Print Sticker")

                        if dialog.exec():
                            if sys.platform.startswith("win"):
                                os.startfile(output, "print")
                            elif sys.platform == "darwin":
                                subprocess.run(["open", "-a", "Preview", output])
                            else:
                                subprocess.run(["xdg-open", output])

                except Exception as e:
                    QMessageBox.warning(self, "Print Error", f"Could not print automatically:\n{e}")



        except Exception as e:
            QMessageBox.critical(self, "❌ Error", str(e))


# ----------------------------------------
# Entry Point
# ----------------------------------------

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StickerApp()
    window.show()
    sys.exit(app.exec())

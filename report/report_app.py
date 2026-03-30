"""
Report Generator Module - PyQt6 UI
Loads an RQS document, extracts fields, lets user review/edit, and generates report.
"""

import os
import sys
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QGroupBox,
    QLineEdit, QScrollArea, QHeaderView, QTextEdit, QSplitter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QColor

from report.rqs_parser import parse_rqs
from report.report_generator import generate_report


class ReportApp(QWidget):
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.rqs_path = None
        self.template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Template.docx")
        self.rqs_data = {}
        self.setWindowTitle("Report Generator")
        self.setGeometry(150, 50, 1100, 850)
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5;
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 13px;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #c0c0c0;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 18px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                padding: 0 8px;
                color: #1a1a2e;
            }
            QPushButton {
                background-color: #1a1a2e;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #16213e;
            }
            QPushButton:disabled {
                background-color: #999;
            }
            QPushButton#btnGenerate {
                background-color: #0f9d58;
                font-size: 15px;
                padding: 14px 30px;
            }
            QPushButton#btnGenerate:hover {
                background-color: #0b7a45;
            }
            QLineEdit, QTextEdit {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 6px;
                background-color: white;
            }
            QTableWidget {
                background-color: white;
                gridline-color: #e0e0e0;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            QHeaderView::section {
                background-color: #1a1a2e;
                color: white;
                padding: 6px;
                border: none;
                font-weight: bold;
            }
        """)

        self._build_ui()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)

        # --- Title ---
        title = QLabel("Report Generator")
        title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #1a1a2e; padding: 8px;")
        main_layout.addWidget(title)

        subtitle = QLabel("RQS (Requirement Sheet) -> Otomatik Rapor")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #666; font-size: 12px; padding-bottom: 6px;")
        main_layout.addWidget(subtitle)

        # --- File Selection ---
        file_group = QGroupBox("Dosya Secimi")
        file_layout = QVBoxLayout(file_group)

        # RQS file
        rqs_row = QHBoxLayout()
        rqs_row.addWidget(QLabel("RQS Dosyasi:"))
        self.lbl_rqs = QLabel("Secilmedi")
        self.lbl_rqs.setStyleSheet("color: #888; font-style: italic;")
        rqs_row.addWidget(self.lbl_rqs, 1)
        btn_rqs = QPushButton("RQS Yukle")
        btn_rqs.clicked.connect(self.load_rqs)
        rqs_row.addWidget(btn_rqs)
        file_layout.addLayout(rqs_row)

        # Template file (fixed path, no selection needed)
        tmpl_row = QHBoxLayout()
        tmpl_row.addWidget(QLabel("Taslak (Template):"))
        self.lbl_template = QLabel("Template.docx (Otomatik)")
        self.lbl_template.setStyleSheet("color: #0f9d58; font-weight: bold;")
        tmpl_row.addWidget(self.lbl_template, 1)
        file_layout.addLayout(tmpl_row)

        main_layout.addWidget(file_group)

        # --- Extracted Data Table ---
        data_group = QGroupBox("RQS'den Cekilen Veriler (Duzenlenebilir)")
        data_layout = QVBoxLayout(data_group)

        self.data_table = QTableWidget()
        self.data_table.setColumnCount(2)
        self.data_table.setHorizontalHeaderLabels(["Alan", "Deger"])
        self.data_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.data_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f8f9fa;
            }
        """)
        data_layout.addWidget(self.data_table)

        main_layout.addWidget(data_group, 1)

        # --- Custom Replacements ---
        custom_group = QGroupBox("Ek Degisiklikler (Opsiyonel)")
        custom_layout = QVBoxLayout(custom_group)

        custom_info = QLabel("Taslakta bulmak istediginiz metni ve yerine koyacaginiz metni yazin:")
        custom_info.setStyleSheet("color: #666; font-size: 11px;")
        custom_layout.addWidget(custom_info)

        self.custom_table = QTableWidget()
        self.custom_table.setColumnCount(2)
        self.custom_table.setHorizontalHeaderLabels(["Bul (Eski Metin)", "Degistir (Yeni Metin)"])
        self.custom_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.custom_table.setRowCount(3)
        self.custom_table.setMaximumHeight(120)
        for i in range(3):
            self.custom_table.setItem(i, 0, QTableWidgetItem(""))
            self.custom_table.setItem(i, 1, QTableWidgetItem(""))
        custom_layout.addWidget(self.custom_table)

        main_layout.addWidget(custom_group)

        # --- Generate Button ---
        btn_layout = QHBoxLayout()

        self.btn_generate = QPushButton("Raporu Olustur")
        self.btn_generate.setObjectName("btnGenerate")
        self.btn_generate.setEnabled(False)
        self.btn_generate.clicked.connect(self.generate)
        btn_layout.addWidget(self.btn_generate)

        btn_back = QPushButton("Geri")
        btn_back.clicked.connect(self.go_back)
        btn_layout.addWidget(btn_back)

        main_layout.addLayout(btn_layout)

        # --- Status ---
        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_status.setStyleSheet("color: #666; font-size: 11px; padding: 4px;")
        main_layout.addWidget(self.lbl_status)

        # --- Author ---
        lbl_author = QLabel("Created by Efe Nakci")
        lbl_author.setAlignment(Qt.AlignmentFlag.AlignRight)
        lbl_author.setStyleSheet("color: gray; font-style: italic; font-size: 11px;")
        main_layout.addWidget(lbl_author)

    def closeEvent(self, event):
        if self.main_window is not None:
            self.main_window.show()
        event.accept()

    def go_back(self):
        if self.main_window is not None:
            self.main_window.show()
        self.close()

    def load_rqs(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "RQS Dosyasi Sec", "",
            "Word Files (*.docx);;All Files (*)"
        )
        if not path:
            return

        self.rqs_path = path
        self.lbl_rqs.setText(os.path.basename(path))
        self.lbl_rqs.setStyleSheet("color: #0f9d58; font-weight: bold;")
        self.lbl_status.setText("RQS okunuyor...")

        try:
            self.rqs_data = parse_rqs(path)
            self._populate_table()
            self.lbl_status.setText(f"RQS basariyla okundu - {len(self.rqs_data) - 1} alan cikarildi")
            self._update_generate_button()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"RQS okunamadi:\n{str(e)}")
            self.lbl_status.setText("RQS okuma hatasi!")

    def _update_generate_button(self):
        self.btn_generate.setEnabled(
            self.rqs_path is not None and
            len(self.rqs_data) > 0
        )

    def _populate_table(self):
        """Fill the data table with extracted RQS fields."""
        # Fields to display (human-readable names)
        display_fields = [
            ("project_no", "Project No."),
            ("task_no", "Task No."),
            ("project", "Project"),
            ("customer", "Customer"),
            ("test_coordinator", "Test Coordinator"),
            ("representative", "Representative"),
            ("component", "Component"),
            ("sample_id", "Sample ID"),
            ("sample_content", "Sample Content"),
            ("start_date", "Start Date"),
            ("test_regulation", "Test Regulation"),
            ("test_object", "Test Object"),
            ("test_fixture", "Test Fixture"),
            ("sled_pulse", "Sled Pulse"),
            ("pulse_id", "Pulse ID"),
            ("direction_of_acceleration", "Direction of Acceleration"),
            ("type_of_dummy", "Type of Dummy"),
            ("camera_setup", "Camera Setup"),
            ("seat_position", "Seat Position"),
            ("seat_back_angle", "Seat Back Angle"),
            ("seat_cushion_angle", "Seat Cushion Angle"),
            ("head_restraint", "Head Restraint"),
            ("test_setup", "Test Setup"),
            ("h_point_x_target", "H-Point X Target"),
            ("h_point_y_target", "H-Point Y Target"),
            ("h_point_z_target", "H-Point Z Target"),
        ]

        # Filter to only existing fields
        existing = [(key, label) for key, label in display_fields if key in self.rqs_data]

        self.data_table.setRowCount(len(existing))
        self._field_keys = []

        for i, (key, label) in enumerate(existing):
            val = self.rqs_data.get(key, "")
            if isinstance(val, list):
                val = "\n".join(val)

            # Field name (read-only)
            name_item = QTableWidgetItem(label)
            name_item.setFlags(name_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            name_item.setBackground(QColor("#f0f0f0"))
            name_item.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            self.data_table.setItem(i, 0, name_item)

            # Value (editable)
            val_item = QTableWidgetItem(str(val))
            self.data_table.setItem(i, 1, val_item)

            self._field_keys.append(key)

    def _get_edited_data(self):
        """Get the (possibly edited) data from the table."""
        data = dict(self.rqs_data)
        for i, key in enumerate(self._field_keys):
            item = self.data_table.item(i, 1)
            if item:
                data[key] = item.text()
        return data

    def _get_custom_replacements(self):
        """Get custom find-and-replace pairs."""
        replacements = {}
        for i in range(self.custom_table.rowCount()):
            old_item = self.custom_table.item(i, 0)
            new_item = self.custom_table.item(i, 1)
            if old_item and new_item:
                old_text = old_item.text().strip()
                new_text = new_item.text().strip()
                if old_text and new_text:
                    replacements[old_text] = new_text
        return replacements

    def generate(self):
        """Generate the report."""
        output_path, _ = QFileDialog.getSaveFileName(
            self, "Raporu Kaydet", "",
            "Word Files (*.docx);;All Files (*)"
        )
        if not output_path:
            return

        if not output_path.endswith(".docx"):
            output_path += ".docx"

        try:
            self.lbl_status.setText("Rapor olusturuluyor...")
            QApplication.processEvents()

            edited_data = self._get_edited_data()
            custom_replacements = self._get_custom_replacements()

            result = generate_report(
                self.template_path,
                edited_data,
                output_path,
                replacements=custom_replacements
            )

            # Count successes
            success_count = sum(1 for v in result.values() if v["success"])
            total_count = len(result)

            self.lbl_status.setText(
                f"Rapor olusturuldu! ({success_count}/{total_count} alan degistirildi)"
            )

            QMessageBox.information(
                self, "Basarili",
                f"Rapor basariyla olusturuldu:\n{output_path}\n\n"
                f"{success_count}/{total_count} alan taslakta bulunup degistirildi."
            )

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Rapor olusturulamadi:\n{str(e)}")
            self.lbl_status.setText("Rapor olusturma hatasi!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ReportApp()
    window.show()
    sys.exit(app.exec())

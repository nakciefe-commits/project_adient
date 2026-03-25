import sys
import os
import subprocess

def _check_and_install_dependencies():
    required_packages = ['pandas', 'numpy', 'PyQt6', 'matplotlib', 'docxtpl', 'openpyxl', 'xlrd']
    for pkg in required_packages:
        try:
            __import__(pkg)
        except ImportError:
            print(f"Eksik kütüphane tespit edildi, yükleniyor: {pkg}...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
                print(f"{pkg} başarıyla yüklendi.")
            except Exception as e:
                print(f"{pkg} yüklenirken hata oluştu: {e}")

_check_and_install_dependencies()

import pandas as pd
import numpy as np
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QFileDialog,
                             QMessageBox, QDoubleSpinBox, QGroupBox, QLineEdit,
                             QCheckBox)
from PyQt6.QtCore import Qt

import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import shared.global_data as global_data
import matplotlib.ticker as mticker

G = 9.81  # g -> m/s² dönüşüm sabiti

# Profesyonel grafik stili
matplotlib.rcParams.update({
    'font.family': 'Arial',
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.titleweight': 'bold',
    'axes.labelsize': 12,
    'axes.labelweight': 'bold',
    'axes.linewidth': 1.2,
    'axes.grid': True,
    'axes.grid.which': 'both',
    'grid.alpha': 0.3,
    'grid.linewidth': 0.8,
    'grid.color': '#b0b0b0',
    'lines.linewidth': 2.2,
    'lines.antialiased': True,
    'xtick.labelsize': 10,
    'ytick.labelsize': 10,
    'xtick.direction': 'in',
    'ytick.direction': 'in',
    'xtick.major.size': 5,
    'ytick.major.size': 5,
    'xtick.minor.visible': True,
    'ytick.minor.visible': True,
    'xtick.minor.size': 3,
    'ytick.minor.size': 3,
    'legend.fontsize': 11,
    'legend.frameon': True,
    'legend.edgecolor': '#cccccc',
    'legend.fancybox': True,
    'legend.shadow': False,
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'savefig.dpi': 300,
    'savefig.bbox': 'tight',
})

class SledAnalyzerApp(QMainWindow):
    def __init__(self, main_window=None):
        super().__init__()
        self.main_window = main_window
        self.setWindowTitle("Sled Test Analyzer (Multi-Graph)")
        self.resize(1100, 900)

        self.excel_path = None
        self.df_actual = None
        self.df_target = None

        # State
        self.current_graph_idx = 0
        self.graphs = ["Spul", "Acceleration vs Velocity", "Actual vs Target Acceleration"]
        self.actual_offset_ms = 0.0
        self.target_offset_ms = 0.0

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)

        # --- Top Area Layout ---
        top_layout = QHBoxLayout()

        # --- Control Panel (Left) ---
        control_group = QGroupBox("Veri Yükleme")
        control_layout = QVBoxLayout()
        control_group.setLayout(control_layout)

        # Single File Selection
        self.btn_excel = QPushButton("Excel Yükle")
        self.btn_excel.setStyleSheet("font-weight: bold; padding: 8px;")
        self.btn_excel.clicked.connect(self.load_excel)
        self.lbl_excel = QLabel("Seçilmedi")
        control_layout.addWidget(self.btn_excel)
        control_layout.addWidget(self.lbl_excel)

        control_layout.addStretch()

        # Action Button
        self.btn_generate = QPushButton("Oluştur / Güncelle")
        self.btn_generate.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.btn_generate.clicked.connect(self.generate_plots)
        control_layout.addWidget(self.btn_generate)

        top_layout.addWidget(control_group, stretch=1)

        # --- Offset Panel (Right) ---
        offset_group = QGroupBox("Offset Ayarları (ms)")
        offset_layout = QVBoxLayout()
        offset_group.setLayout(offset_layout)

        # Actual Offset
        actual_row = QHBoxLayout()
        actual_row.addWidget(QLabel("Actual Offset (ms):"))
        self.spin_actual_offset = QDoubleSpinBox()
        self.spin_actual_offset.setRange(-10000.0, 10000.0)
        self.spin_actual_offset.setValue(0.0)
        self.spin_actual_offset.setSingleStep(0.4)
        self.spin_actual_offset.setDecimals(1)
        self.spin_actual_offset.valueChanged.connect(self.on_actual_offset_changed)
        actual_row.addWidget(self.spin_actual_offset)
        offset_layout.addLayout(actual_row)

        # Target Offset
        target_row = QHBoxLayout()
        target_row.addWidget(QLabel("Target Offset (ms):"))
        self.spin_target_offset = QDoubleSpinBox()
        self.spin_target_offset.setRange(-10000.0, 10000.0)
        self.spin_target_offset.setValue(0.0)
        self.spin_target_offset.setSingleStep(0.4)
        self.spin_target_offset.setDecimals(1)
        self.spin_target_offset.valueChanged.connect(self.on_target_offset_changed)
        target_row.addWidget(self.spin_target_offset)
        offset_layout.addLayout(target_row)

        # 14 ms checkbox
        self.check_14ms = QCheckBox("Evrensel 14 ms Offset (her ikisine de)")
        self.check_14ms.stateChanged.connect(self.apply_14ms_offset)
        offset_layout.addWidget(self.check_14ms)

        offset_layout.addStretch()
        top_layout.addWidget(offset_group, stretch=2)

        main_layout.addLayout(top_layout)

        # --- Graph Navigation ---
        nav_layout = QHBoxLayout()
        self.btn_prev = QPushButton("⬅")
        self.btn_prev.setStyleSheet("font-size: 24px; font-weight: bold; width: 60px; height: 40px;")
        self.btn_prev.clicked.connect(self.prev_graph)

        self.lbl_graph_name = QLabel(f"{self.graphs[self.current_graph_idx]}")
        self.lbl_graph_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_graph_name.setStyleSheet("font-size: 16px; font-weight: bold;")

        self.btn_next = QPushButton("➡")
        self.btn_next.setStyleSheet("font-size: 24px; font-weight: bold; width: 60px; height: 40px;")
        self.btn_next.clicked.connect(self.next_graph)

        nav_layout.addWidget(self.btn_prev)
        nav_layout.addWidget(self.lbl_graph_name)
        nav_layout.addWidget(self.btn_next)

        main_layout.addLayout(nav_layout)

        # --- Plot Area (Matplotlib) ---
        plot_group = QGroupBox("Grafik Ekranı")
        plot_layout = QVBoxLayout()
        plot_group.setLayout(plot_layout)

        self.figure = Figure(figsize=(10, 8))
        self.canvas = FigureCanvas(self.figure)
        plot_layout.addWidget(self.canvas)

        import matplotlib.gridspec as gridspec
        self.gs = gridspec.GridSpec(2, 1, height_ratios=[4, 1.2])
        self.ax = self.figure.add_subplot(self.gs[0])
        self.ax_table = self.figure.add_subplot(self.gs[1])
        self.ax_table.axis('off')

        self.ax2 = None

        main_layout.addWidget(plot_group)

        # --- Export Area ---
        export_layout = QHBoxLayout()
        export_layout.addWidget(QLabel("Kayıt Dizini:"))
        self.txt_export = QLineEdit(r"c:\Users\pc1\Desktop\adient_data\velocity_acc_target_spul")
        export_layout.addWidget(self.txt_export)

        self.btn_browse = QPushButton("Gözat...")
        self.btn_browse.clicked.connect(self.browse_export_dir)
        export_layout.addWidget(self.btn_browse)

        self.btn_export = QPushButton("Tüm Grafikleri Kaydet (.png)")
        self.btn_export.clicked.connect(self.export_plots)

        self.btn_report = QPushButton("Rapor Oluştur (Word)")
        self.btn_report.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")
        self.btn_report.clicked.connect(self.generate_word_report)

        self.btn_back = QPushButton("Ana Menüye Dön")
        self.btn_back.setStyleSheet("background-color: #9E9E9E; color: white; font-weight: bold;")
        self.btn_back.clicked.connect(self.close)

        export_layout.addWidget(self.btn_export)
        export_layout.addWidget(self.btn_report)
        export_layout.addWidget(self.btn_back)

        main_layout.addLayout(export_layout)

        # --- Author Info ---
        lbl_author = QLabel("Created by Efe Nakcı")
        lbl_author.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        lbl_author.setStyleSheet("color: gray; font-style: italic; font-size: 11px; padding-top: 5px;")
        main_layout.addWidget(lbl_author)

    def closeEvent(self, event):
        if self.main_window is not None:
            self.main_window.show()
        event.accept()

    def on_actual_offset_changed(self, val):
        self.actual_offset_ms = val
        if self.df_actual is not None:
            self.draw_current_graph()

    def on_target_offset_changed(self, val):
        self.target_offset_ms = val
        if self.df_actual is not None:
            self.draw_current_graph()

    def apply_14ms_offset(self, state):
        if state == Qt.CheckState.Checked.value:
            self.spin_actual_offset.setValue(14.0)
            self.spin_target_offset.setValue(14.0)
            self.spin_actual_offset.setEnabled(False)
            self.spin_target_offset.setEnabled(False)
        else:
            self.spin_actual_offset.setEnabled(True)
            self.spin_target_offset.setEnabled(True)

    def browse_export_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "Kayıt Klasörü Seç", self.txt_export.text())
        if directory:
            self.txt_export.setText(directory)

    def prev_graph(self):
        self.current_graph_idx = (self.current_graph_idx - 1) % len(self.graphs)
        self.update_graph_view()

    def next_graph(self):
        self.current_graph_idx = (self.current_graph_idx + 1) % len(self.graphs)
        self.update_graph_view()

    def update_graph_view(self):
        self.lbl_graph_name.setText(f"{self.graphs[self.current_graph_idx]}")
        if self.df_actual is not None:
            self.draw_current_graph()

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seç", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.excel_path = path
            self.lbl_excel.setText(os.path.basename(path))

    def generate_plots(self):
        if not self.excel_path:
            QMessageBox.warning(self, "Uyarı", "Lütfen Excel dosyası yükleyin.")
            return

        try:
            # Actual data: A(0)=Time, F(5)=Acceleration, G(6)=Velocity
            df_actual_raw = pd.read_excel(self.excel_path, skiprows=9, usecols=[0, 5, 6])
            df_actual_raw.columns = ['Time', 'Acceleration', 'Velocity']

            # Sayısal dönüşüm
            for col in df_actual_raw.columns:
                df_actual_raw[col] = pd.to_numeric(df_actual_raw[col], errors='coerce')

            # g -> m/s² dönüşümü
            df_actual_raw['Acceleration'] = df_actual_raw['Acceleration'] * G

            # Spul hesapla: V² / t
            df_actual_raw['Spul'] = np.where(
                (df_actual_raw['Time'] != 0) & (df_actual_raw['Time'].notna()),
                (df_actual_raw['Velocity'] ** 2) / df_actual_raw['Time'],
                0
            )
            self.df_actual = df_actual_raw

            # Target data: A(0)=Time, B(1)=Target Acceleration, C(2)=Target Velocity
            df_target_raw = pd.read_excel(self.excel_path, skiprows=9, usecols=[0, 1, 2])
            df_target_raw.columns = ['Time', 'Target Acceleration', 'Target Velocity']

            for col in df_target_raw.columns:
                df_target_raw[col] = pd.to_numeric(df_target_raw[col], errors='coerce')

            # g -> m/s² dönüşümü
            df_target_raw['Target Acceleration'] = df_target_raw['Target Acceleration'] * G

            # Target Spul hesapla
            df_target_raw['Spul'] = np.where(
                (df_target_raw['Time'] != 0) & (df_target_raw['Time'].notna()),
                (df_target_raw['Target Velocity'] ** 2) / df_target_raw['Time'],
                0
            )
            self.df_target = df_target_raw

            self.draw_current_graph()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Veri işlenirken bir hata oluştu:\n{str(e)}")

    def apply_offset(self, df, offset_ms):
        df_plot = df.copy()
        offset_sec = offset_ms / 1000.0
        df_plot['Offset_Time'] = df_plot['Time'] - offset_sec
        return df_plot[df_plot['Offset_Time'] >= 0]

    def _cleanup_axes(self):
        self.ax.clear()
        if self.ax2 is not None:
            self.ax2.remove()
            self.ax2 = None
        self.ax_table.clear()
        self.ax_table.axis('off')

    def draw_current_graph(self):
        if self.df_actual is None:
            return

        df_actual_plot = self.apply_offset(self.df_actual, self.actual_offset_ms)
        df_target_plot = self.apply_offset(self.df_target, self.target_offset_ms) if self.df_target is not None else None

        self._cleanup_axes()

        idx = self.current_graph_idx

        if idx == 0:
            self._draw_spul(df_actual_plot, df_target_plot)
        elif idx == 1:
            self._draw_acc_vel(df_actual_plot)
        elif idx == 2:
            self._draw_acc_target_acc(df_actual_plot, df_target_plot)

        self.figure.tight_layout()
        self.canvas.draw()

    def _draw_spul(self, df_plot, df_target_plot):
        if 'Spul' not in df_plot.columns:
            return

        actual_color = '#1F4E79'   # Koyu mavi
        target_color = '#ED7D31'   # Turuncu (Excel tarzı)

        self.ax.plot(df_plot['Offset_Time'].values, df_plot['Spul'].values,
                     color=actual_color, linewidth=2.5, label="SPUL", zorder=3)
        max_actual_spul = df_plot['Spul'].max()
        idx_max = df_plot['Spul'].idxmax()
        if pd.isna(idx_max): max_actual_time_sec = 0
        else: max_actual_time_sec = df_plot.loc[idx_max, 'Offset_Time']

        max_target_spul = "-"
        max_target_time_ms = "-"
        if df_target_plot is not None and 'Spul' in df_target_plot.columns:
            self.ax.plot(df_target_plot['Offset_Time'].values, df_target_plot['Spul'].values,
                         color=target_color, linewidth=2.5, linestyle='-', label="Target Spul", zorder=2)
            max_target_spul = df_target_plot['Spul'].max()
            t_idx = df_target_plot['Spul'].idxmax()
            if not pd.isna(t_idx):
                max_target_time_sec = df_target_plot.loc[t_idx, 'Offset_Time']
                max_target_time_ms = max_target_time_sec * 1000.0

        self.ax.set_title("SPUL", pad=12)
        self.ax.set_xlabel("Time, (s)", labelpad=8)
        self.ax.set_ylabel("Spul Value, (m²/s³)", labelpad=8)
        self.ax.legend(
            loc='upper center',
            bbox_to_anchor=(0.5, -0.12),
            ncol=2,
            frameon=True,
            fontsize=11,
            handlelength=3.0,
            handleheight=1.5,
            columnspacing=4.0,
        )
        self.ax.grid(True, which='major', alpha=0.4, linewidth=0.8)
        self.ax.grid(True, which='minor', alpha=0.15, linewidth=0.5)
        self.ax.minorticks_on()
        self.ax.set_xlim(left=0)
        self.ax.set_ylim(bottom=0)
        self.ax.spines['top'].set_visible(False)
        self.ax.spines['right'].set_visible(False)

        # Tablo
        actual_val_str = f"{max_actual_spul:.1f}  m²/s³   ({max_actual_time_sec*1000.0:.1f} ms)" if not pd.isna(max_actual_spul) else "-"
        target_val_str = f"{max_target_spul:.1f}  m²/s³   ({max_target_time_ms:.1f} ms)" if max_target_spul != "-" else "-"

        cell_text = [
            ["SPUL", actual_val_str, ""],
            ["Target Spul", target_val_str, ""]
        ]
        self._build_table(cell_text, "SPUL\nSpecific Accident Capability\nf(t) = v² / t")

    def _draw_acc_vel(self, df_plot):
        if 'Acceleration' not in df_plot.columns or 'Velocity' not in df_plot.columns:
            self.ax.text(0.5, 0.5, "Acceleration veya Velocity Sütunu Bulunamadı", ha='center', va='center')
            return

        acc_color = '#1F4E79'   # Koyu mavi
        vel_color = '#548235'   # Koyu yeşil

        self.ax2 = self.ax.twinx()

        l1 = self.ax.plot(df_plot['Offset_Time'].values, df_plot['Acceleration'].values,
                          color=acc_color, linewidth=2.5, label="Acceleration", zorder=3)
        l2 = self.ax2.plot(df_plot['Offset_Time'].values, df_plot['Velocity'].values,
                           color=vel_color, linewidth=2.5, label="Velocity", zorder=2)

        max_acc = df_plot['Acceleration'].max()
        a_idx = df_plot['Acceleration'].idxmax()
        max_acc_t = df_plot.loc[a_idx, 'Offset_Time'] if not pd.isna(a_idx) else 0

        max_vel = df_plot['Velocity'].max()
        v_idx = df_plot['Velocity'].idxmax()
        max_vel_t = df_plot.loc[v_idx, 'Offset_Time'] if not pd.isna(v_idx) else 0

        self.ax.set_title("Sled Acceleration and Velocity", pad=12)
        self.ax.set_xlabel("Time, (s)", labelpad=8)
        self.ax.set_ylabel("Acceleration, (m/s²)", labelpad=8, color=acc_color)
        self.ax2.set_ylabel("Velocity, (m/s)", labelpad=8, color=vel_color)
        self.ax.tick_params(axis='y', colors=acc_color)
        self.ax2.tick_params(axis='y', colors=vel_color)

        lines = l1 + l2
        labels = [l.get_label() for l in lines]
        self.ax.legend(
            lines, labels,
            loc='upper center',
            bbox_to_anchor=(0.5, -0.12),
            ncol=2,
            frameon=True,
            fontsize=11,
            handlelength=3.0,
            handleheight=1.5,
            columnspacing=4.0,
        )
        self.ax.grid(True, which='major', alpha=0.4, linewidth=0.8)
        self.ax.grid(True, which='minor', alpha=0.15, linewidth=0.5)
        self.ax.minorticks_on()
        self.ax.set_xlim(left=0)
        self.ax.spines['top'].set_visible(False)

        # Tablo
        v_str = f"{max_vel:.2f} m/s ({max_vel_t*1000.0:.1f} ms)"
        a_str = f"{max_acc:.2f} m/s²     ({max_acc_t*1000.0:.1f} ms)"
        cell_text = [
            ["Sled Velocity", v_str, ""],
            ["Sled Acceleration", a_str, ""]
        ]
        self._build_table(cell_text, "Sled Acceleration and Velocity")

    def _draw_acc_target_acc(self, df_plot, df_target_plot):
        if 'Acceleration' not in df_plot.columns:
            self.ax.text(0.5, 0.5, "Actual'da Acceleration Sütunu Bulunamadı", ha='center', va='center')
            return

        acc_color = '#1F4E79'          # Koyu mavi
        target_pulse_color = '#ED7D31'  # Turuncu

        l1 = self.ax.plot(df_plot['Offset_Time'].values, df_plot['Acceleration'].values,
                          color=acc_color, linewidth=2.5, label="Acceleration", zorder=3)

        max_acc = df_plot['Acceleration'].max()
        a_idx = df_plot['Acceleration'].idxmax()
        max_acc_t = df_plot.loc[a_idx, 'Offset_Time'] if not pd.isna(a_idx) else 0

        max_t_acc = "-"
        max_t_acc_t = "-"
        l2 = []
        if df_target_plot is not None and 'Target Acceleration' in df_target_plot.columns:
            l2 = self.ax.plot(df_target_plot['Offset_Time'].values, df_target_plot['Target Acceleration'].values,
                              color=target_pulse_color, linewidth=2.5, label="Target Pulse", zorder=2)
            max_t_acc = df_target_plot['Target Acceleration'].max()
            ta_idx = df_target_plot['Target Acceleration'].idxmax()
            if not pd.isna(ta_idx):
                max_t_acc_t_sec = df_target_plot.loc[ta_idx, 'Offset_Time']
                max_t_acc_t = max_t_acc_t_sec * 1000.0

        self.ax.set_title("Sled vs. Target Acceleration", pad=12)
        self.ax.set_xlabel("Time, (s)", labelpad=8)
        self.ax.set_ylabel("Acceleration, (m/s²)", labelpad=8)

        lines = l1 + l2
        labels = [l.get_label() for l in lines]
        self.ax.legend(
            lines, labels,
            loc='upper center',
            bbox_to_anchor=(0.5, -0.12),
            ncol=2,
            frameon=True,
            fontsize=11,
            handlelength=3.0,
            handleheight=1.5,
            columnspacing=4.0,
        )
        self.ax.grid(True, which='major', alpha=0.4, linewidth=0.8)
        self.ax.grid(True, which='minor', alpha=0.15, linewidth=0.5)
        self.ax.minorticks_on()
        self.ax.set_xlim(left=0)
        self.ax.spines['top'].set_visible(False)
        self.ax.spines['right'].set_visible(False)

        # Tablo
        a_str = f"{max_acc:.2f} m/s²     ({max_acc_t*1000.0:.1f} ms)"
        t_str = f"{max_t_acc:.2f} m/s²     ({max_t_acc_t:.1f} ms)" if max_t_acc != "-" else "-"
        cell_text = [
            ["Sled Acceleration", a_str, ""],
            ["Target Acceleration", t_str, ""]
        ]
        self._build_table(cell_text, "Sled vs. Target Acceleration")

    def _build_table(self, cell_text, graph_name_text):
        col_labels = ["", "Max. Value", "Graph Name"]
        table = self.ax_table.table(cellText=cell_text, colLabels=col_labels, loc='center', cellLoc='center', bbox=[0, 0, 1, 1])
        table.auto_set_font_size(False)
        table.set_fontsize(10)

        for (row, col), cell in table.get_celld().items():
            cell.set_text_props(ha='center', va='center')
            if row == 0:
                cell.set_text_props(weight='bold', ha='center', va='center')

            if col == 2 and row == 2:
                cell.visible_edges = 'BRL'
            if col == 2 and row == 1:
                cell.visible_edges = 'TRL'

        self.ax_table.text(0.833, 0.333, graph_name_text, ha='center', va='center', fontsize=10, transform=self.ax_table.transAxes)

    def export_plots(self):
        save_dir = self.txt_export.text()
        if not os.path.exists(save_dir) or not os.path.isdir(save_dir):
            QMessageBox.warning(self, "Hata", "Geçersiz kayıt dizini.")
            return

        if self.df_actual is None:
            QMessageBox.warning(self, "Hata", "İşlenecek veri yok!")
            return

        try:
            saved_idx = self.current_graph_idx

            names = ["Spul.png", "Acc_vs_Vel.png", "Acc_vs_Targetacc.png"]

            for i in range(3):
                self.current_graph_idx = i
                self.draw_current_graph()
                path = os.path.join(save_dir, names[i])
                self.figure.savefig(path, dpi=300, bbox_inches='tight')

            # Restore
            self.current_graph_idx = saved_idx
            self.update_graph_view()

            QMessageBox.information(self, "Başarılı", f"Tüm 3 grafik seçilen klasöre kaydedildi:\n{names[0]}, {names[1]}, {names[2]}")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Dışa aktarma hatası:\n{str(e)}")

    def generate_word_report(self):
        save_dir = self.txt_export.text()
        if not os.path.exists(save_dir) or not os.path.isdir(save_dir):
            QMessageBox.warning(self, "Hata", "Geçersiz dizin.")
            return

        if self.df_actual is None:
            QMessageBox.warning(self, "Hata", "İşlenecek veri yok!")
            return

        template_path = os.path.join(save_dir, "Template.docx")
        if not os.path.exists(template_path):
            QMessageBox.warning(self, "Hata", f"Template.docx dosyası bulunamadı, aynı dizinde olmalı:\n{template_path}")
            return

        test_no_input = global_data.config.get("TEST_NO", "Belirtilmedi")

        suffix = test_no_input.split('/')[-1] if '/' in test_no_input else test_no_input
        out_filename = f"graphs_{suffix}.docx"

        try:
            import tempfile
            temp_dir = tempfile.mkdtemp()

            saved_idx = self.current_graph_idx
            paths = {}
            labels = ["Spul", "Acc_vs_Vel", "Acc_vs_Targetacc"]

            for i in range(3):
                self.current_graph_idx = i
                self.draw_current_graph()
                path = os.path.join(temp_dir, f"{labels[i]}.png")
                self.figure.savefig(path, dpi=300, bbox_inches='tight')
                paths[labels[i]] = path

            self.current_graph_idx = saved_idx
            self.update_graph_view()

            doc = DocxTemplate(template_path)

            context = {
                "TEST_NO": global_data.config.get("TEST_NO", ""),
                "TEST_DATE": global_data.config.get("TEST_DATE", ""),
                "PROJECT": global_data.config.get("PROJECT", ""),
                "SPUL": InlineImage(doc, paths["Spul"], width=Mm(160)),
                "ACC_VEL": InlineImage(doc, paths["Acc_vs_Vel"], width=Mm(160)),
                "ACC_TARGET": InlineImage(doc, paths["Acc_vs_Targetacc"], width=Mm(160))
            }

            doc.render(context)

            root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            tempfiles_dir = os.path.join(root_dir, "tempfiles")
            if not os.path.exists(tempfiles_dir):
                os.makedirs(tempfiles_dir)

            final_out_path = os.path.join(tempfiles_dir, out_filename)
            doc.save(final_out_path)

            QMessageBox.information(self, "Başarılı", f"Word Raporu başarıyla oluşturuldu!\n\nDosya Yolu: {final_out_path}")

            if self.main_window:
                self.close()
                self.main_window.show()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Rapor oluşturulurken hata oluştu:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SledAnalyzerApp()
    window.show()
    sys.exit(app.exec())

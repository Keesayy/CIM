# Cal Info Mesure GUI - 9020B
# By Arthur Péraud 12/2025
import sys
import os
import openpyxl
import pandas as pd
import numpy as np
import time
import pyvisa
from pathlib import Path
from openpyxl.workbook import Workbook

from PySide6.QtWidgets import (
    QApplication, QWidget, QFormLayout, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QMessageBox, QPlainTextEdit, QGridLayout
)
from PySide6.QtCore import Qt, QCoreApplication, QThread, Signal
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QLabel

### Constants
PRECISION = 3
NB_DIGIT = 3      # Power Meter
NB_ESE_BITS = 60
NB_QUES_BITS = 952
DEBUG = False

def PRINT(*args, **kwargs):
    if DEBUG:
        return __builtins__.print(*args, **kwargs)

def Hz_to_GHz(x: float) -> float:
    return x * 1E9

def Float_precision_str(n: int) -> str:
    s = '#,##0.'
    for i in range(n):
        s += '0'
    return s

def Excel_name(name: str, precision: int, freq_start: float, freq_stop: float, nb_points: float, 
               dwel: float, amp_start: float, amp_stop: float, amp_inc: float, freq_multi : float) -> str:

    def format_amp(amp: float) -> str:
        if amp < 0:
            return f"Neg{int(abs(amp))}"
        return str(int(amp))

    s = name
    if freq_multi > 1:
        s += f"_MI-757-{int(freq_multi)}X.xlsx"
    else:
        s += f"_{format_amp(amp_start)}"
        s += f"_{format_amp(amp_stop)}"
        s += f'_MI-9020B.xlsx'
    return Path(s)

def Excel_Index(n):
    if n < 0: raise ValueError("n<0")
    result = []
    n += 1
    while n > 0:
        n -= 1
        result.append(chr(ord('A') + (n % 26)))
        n //= 26
    return ''.join(reversed(result))

def Get_unique_filename(path: str) -> str:
    base, ext = os.path.splitext(path)
    i = 1
    new_path = path
    while os.path.exists(new_path):
        new_path = f"{base}({i}){ext}"
        i += 1
    return new_path

def Build_path_names(name: str, client: str, year: int) -> Path:
    base_dir = Path(rf"E:\\Cal Info Mesure\\{client}\\Data {year}")
    base_dir.mkdir(parents=True, exist_ok=True)
    return base_dir / name

def Save_workbook_safely(wb: Workbook, output_file: str, log_func=None) -> None:
    if os.path.exists(output_file):
        if log_func:
            log_func(f"⚠️ Fichier existe: {output_file}")
        new_file = Get_unique_filename(output_file)
        wb.save(new_file)
        if log_func:
            log_func(f"📁 Sauvegardé: {new_file}")
    else:
        wb.save(output_file)
        if log_func:
            log_func(f"✅ Sauvegardé: {output_file}")

class AcquisitionThread(QThread):
    log_signal = Signal(str)
    finished_signal = Signal(str)
    error_signal = Signal(str)
    
    def __init__(self, year, client, freq_start, freq_stop, nb_points_swf, dwel, amp_start, amp_inc, amp_stop, freq_multi):
        super().__init__()
        self.year = year
        self.client = client
        self.freq_start = freq_start
        self.freq_stop = freq_stop
        self.nb_points_swf = nb_points_swf
        self.dwel = dwel
        self.amp_start = amp_start
        self.amp_inc = amp_inc
        self.amp_stop = amp_stop
        self.freq_multi = freq_multi
        
    def log(self, message):
        self.log_signal.emit(message)
    
    def run(self):
        try:
            rm, power_meter, signal_source = Gpid_devices_open()
            
            Signal_source_init(signal_source)
            Power_meter_init(power_meter)
            
            excel_name = Excel_name('Flatness', PRECISION, self.freq_start, self.freq_stop, 
                                  self.nb_points_swf, self.dwel, self.amp_start, self.amp_stop, self.amp_inc, self.freq_multi)
            
            excel = openpyxl.Workbook()
            sheet = excel.active
            sheet.title = "Data"
            sheet['B1'] = 'Fréquence (GHz)'
            
            self.log('START ACQUISITION')
            start_tot = time.time()
            
            if self.amp_inc == 0:
                col = 'C'
                Sweep_freq(excel, power_meter, signal_source, self.freq_start, self.freq_stop, 
                          self.nb_points_swf, self.dwel, self.amp_start, self.freq_multi, col, self.log)
            else:
                k=2
                for i in range(int(self.amp_start), int(self.amp_stop)+1, int(self.amp_inc)):
                    col = Excel_Index(k)
                    Sweep_freq(excel, power_meter, signal_source, self.freq_start, self.freq_stop, 
                              self.nb_points_swf, self.dwel, i, self.freq_multi, col, self.log)
                    k = k + 1
            
            end_tot = time.time()
            self.log(f'TOTAL TIME: {end_tot - start_tot:.3f}s')
            
            output_file = Build_path_names(str(excel_name), self.client, self.year)
            Save_workbook_safely(excel, str(output_file), self.log)
            
            CLOSE_ALL(signal_source, power_meter, excel, rm)
            self.finished_signal.emit(str(output_file))
            
        except Exception as e:
            self.error_signal.emit(str(e))

def Gpid_devices_open():
    rm = pyvisa.ResourceManager()
    power_meter = rm.open_resource('GPIB0::13::INSTR')
    signal_source = rm.open_resource('TCPIP::192.168.10.100::INSTR')
    
    power_meter.write('SYST:LANG SCPI')
    signal_source.write('SYST:LANG SCPI')
    time.sleep(0.2)
    
    return rm, power_meter, signal_source

def Signal_source_init(signal_source):
    # signal_source.write('*RST')
    signal_source.write('*CLS')
    signal_source.write('*ESE 0')
    signal_source.write('*SRE 0')

def Power_meter_init(power_meter):
    power_meter.write('*CLS')
    power_meter.write('*ESE 1')
    power_meter.write('UNIT:POW dBm')
    power_meter.write(f'DISP:RES {NB_DIGIT}')

def Show_parameters_sweep_freq(freq_start, freq_stop, nb_points, dwel, amplitude, log_func):
    log_func(f'SWEEP FREQ | fstart:{freq_start:.3f}GHz fstop:{freq_stop:.3f}GHz pts:{nb_points} dwel:{dwel:.3f}ms amp:{amplitude}dBm')

def Sweep_freq(excel, power_meter, signal_source, freq_start, freq_stop, freq_step, dwel, amplitude, freq_multi, col, log_func):
    freq_start = Hz_to_GHz(freq_start)
    freq_stop = Hz_to_GHz(freq_stop)
    freq_step = Hz_to_GHz(freq_step)
    
    sheet = excel.active
    sheet[f'{col}1'] = f'{amplitude} dBm'
    
    signal_source.write('OUTP ON')
    signal_source.write(f'POW {amplitude} dBm')
    
    freqcal = freq_start / freq_multi
    freq = freq_start
    precision_string = Float_precision_str(PRECISION)
    
    nb_points = int((freq_stop - freq_start) / freq_step + 1)

    Show_parameters_sweep_freq(freq_start, freq_stop, nb_points, dwel, amplitude, log_func)

    for i in range(nb_points):
        signal_source.write(f'FREQ:CW {freqcal}')
        
        power_meter.write('*CLS')
        power_meter.write(f'FREQ {freq}')
        
        power_meter.write('TRIG:DEL:AUTO ON')
        power_meter.write('INIT:CONT OFF')
        power_meter.write('TRIG:SOUR IMM')
        power_meter.write('INIT')
        power_meter.write('*OPC')
        
        STB_polling(power_meter, signal_source, timeout=20, sleepTime=0.15)
        time.sleep(1)
        
        power_meter.query('*ESR?')
        level = power_meter.query('FETCH?')
        
        log_func(f'{i}: {freq/1e9:.3f}GHz | {float(level):.3f}dBm')
        
        sheet[f'B{i+3}'] = freq / 1e9
        sheet[f'{col}{i+3}'] = float(level)
        sheet[f'B{i+3}'].number_format = precision_string
        sheet[f'{col}{i+3}'].number_format = precision_string
        
        freq += freq_step
        freqcal += freq_step / freq_multi
    

def CLOSE_ALL(signal_source, power_meter, excel, rm):
    signal_source.close()
    power_meter.close()
    excel.close()
    rm.close()

def STB_polling(instrument, instrument_bis, condition=32, timeout=1.0, sleepTime=0.3):
    end_time = time.time() + timeout
    status = False
    error = False
    stb = instrument.read_stb()
    status = (stb & condition) == condition
    error = (stb & 1) == 1
    
    while not status and time.time() < end_time and not error:
        time.sleep(sleepTime)
        stb = instrument.read_stb()
        status = (stb & condition) == condition
        error = (stb & 4) == 4
    
    return status

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Cal Info Mesure – 9020B")
        self.resize(500, 450)  # Augmenté la hauteur pour le nouveau champ

        # Logo
        base_dir = Path(__file__).resolve().parent
        icon_path = base_dir / "Logo_CIM.jpg"
        self.setWindowIcon(QIcon(str(icon_path)))

        # Champs - placeholders gris
        self.year_edit = QLineEdit()
        self.year_edit.setPlaceholderText("2025")
        self.client_edit = QLineEdit()
        self.client_edit.setPlaceholderText("Nom du client")

        self.freq_start_edit = QLineEdit()
        self.freq_start_edit.setPlaceholderText("0.01")
        self.freq_stop_edit = QLineEdit()
        self.freq_stop_edit.setPlaceholderText("20.5")
        self.nb_points_swf_edit = QLineEdit()
        self.nb_points_swf_edit.setPlaceholderText("1")
        self.dwel_edit = QLineEdit()
        self.dwel_edit.setPlaceholderText("1.00")
        self.amp_start_edit = QLineEdit()
        self.amp_start_edit.setPlaceholderText("-30")
        self.amp_inc_edit = QLineEdit()
        self.amp_inc_edit.setPlaceholderText("1")
        self.amp_stop_edit = QLineEdit()
        self.amp_stop_edit.setPlaceholderText("15")       
        self.freq_multi_edit = QLineEdit()
        self.freq_multi_edit.setPlaceholderText("1")

        # Grille
        grid = QGridLayout()
        grid.addWidget(QLabel("Freq start (GHz):"), 0, 0)
        grid.addWidget(self.freq_start_edit, 0, 1)
        grid.addWidget(QLabel("Freq stop (GHz):"), 0, 2)
        grid.addWidget(self.freq_stop_edit, 0, 3)

        grid.addWidget(QLabel("Freq inc (GHz):"), 1, 0)
        grid.addWidget(self.nb_points_swf_edit, 1, 1)
        grid.addWidget(QLabel("Dwel (ms):"), 1, 2)
        grid.addWidget(self.dwel_edit, 1, 3)

        grid.addWidget(QLabel("Amp start (dBm):"), 2, 0)
        grid.addWidget(self.amp_start_edit, 2, 1)
        grid.addWidget(QLabel("Amp inc (dBm):"), 2, 2)
        grid.addWidget(self.amp_inc_edit, 2, 3)

        grid.addWidget(QLabel("Amp stop (dBm):"), 3, 0)
        grid.addWidget(self.amp_stop_edit, 3, 1)
        
        grid.addWidget(QLabel("Freq-Multi:"), 3, 2)
        grid.addWidget(self.freq_multi_edit, 3, 3)

        form = QFormLayout()
        form.addRow("Année :", self.year_edit)
        form.addRow("Client :", self.client_edit)

        self.run_button = QPushButton("START ACQUISITION")
        self.run_button.clicked.connect(self.on_ok_clicked)
        self.exit_button = QPushButton("Quitter")
        self.exit_button.clicked.connect(self.close)

        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.run_button)
        buttons_layout.addWidget(self.exit_button)

        logo_label = QLabel()
        logo_pixmap = QPixmap(str(icon_path))
        logo_pixmap = logo_pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(logo_pixmap)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)

        layout = QVBoxLayout()
        layout.addLayout(form)
        layout.addLayout(grid)
        layout.addWidget(logo_label)
        layout.addLayout(buttons_layout)
        layout.addWidget(self.log_edit)

        self.setLayout(layout)
        
        self.acquisition_thread = None

    def log(self, message: str):
        self.log_edit.appendPlainText(message)
        QCoreApplication.processEvents()

    def on_ok_clicked(self):
        year_text = self.year_edit.text().strip()
        client = self.client_edit.text().strip()

        if not year_text or not year_text.isdigit():
            QMessageBox.warning(self, "Erreur", "Année : entier requis (2025)")
            return
        if not client:
            QMessageBox.warning(self, "Erreur", "Client requis")
            return

        year = int(year_text)
        
        try:
            freq_start = float(self.freq_start_edit.text() or "0.01")
            freq_stop = float(self.freq_stop_edit.text() or "20.5")
            nb_points_swf = float(self.nb_points_swf_edit.text() or "1")
            dwel = float(self.dwel_edit.text() or "1.00")
            amp_start = float(self.amp_start_edit.text() or "-30")
            amp_inc = float(self.amp_inc_edit.text() or "1")
            amp_stop = float(self.amp_stop_edit.text() or "15")
            freq_multi = float(self.freq_multi_edit.text() or "1")

            self.log("=== DÉMARRAGE ACQUISITION ===")
            self.run_button.setEnabled(False)
            
            self.acquisition_thread = AcquisitionThread(
                year, client, freq_start, freq_stop, nb_points_swf, dwel, 
                amp_start, amp_inc, amp_stop, freq_multi
            )
            self.acquisition_thread.log_signal.connect(self.log)
            self.acquisition_thread.finished_signal.connect(self.on_acquisition_finished)
            self.acquisition_thread.error_signal.connect(self.on_acquisition_error)
            self.acquisition_thread.start()
            
        except ValueError as e:
            QMessageBox.warning(self, "Erreur saisie", f"Float/Int invalide:\n{e}")

    def on_acquisition_finished(self, output_file):
        self.log(f"✅ ACQUISITION TERMINÉE: {output_file}")
        self.run_button.setEnabled(True)
        QMessageBox.information(self, "Succès", f"Fichier généré:\n{output_file}")

    def on_acquisition_error(self, error_msg):
        self.log(f"❌ ERREUR: {error_msg}")
        self.run_button.setEnabled(True)
        QMessageBox.critical(self, "Erreur", error_msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    base_dir = Path(__file__).resolve().parent
    icon_path = base_dir / "Logo_CIM.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

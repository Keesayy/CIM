# Pulse Generator GUI - BNC505 by Arthur Péraud 17/02/2026

import sys
import time
from pathlib import Path

import pyvisa

from PySide6.QtWidgets import (
    QApplication, QWidget, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QMessageBox, QPlainTextEdit, QGridLayout,
    QLabel, QCheckBox, QComboBox, QSizePolicy
)
from PySide6.QtCore import Qt, QCoreApplication, QThread, Signal, QDateTime
from PySide6.QtGui import QIcon, QPixmap

PRECISION = 2

def MHz_to_s(mhz: float) -> float:
    if mhz <= 0:
        raise ValueError("Frequence doit être > 0 MHz")
    return 1.0 / (mhz * 1e6)


def Gpid_devices_open():
    rm = pyvisa.ResourceManager()
    pulse_generator = rm.open_resource("GPIB0::9::INSTR")
    return rm, pulse_generator


def Pulse_generator_init(pulse_generator) -> None:
    pulse_generator.write("*RST")
    pulse_generator.write("*CLS")


def CLOSE_ALL(pulse_generator, rm) -> None:
    try:
        pulse_generator.close()
    except Exception:
        pass
    try:
        rm.close()
    except Exception:
        pass


def Create_pulse(pulse_generator,
                 channel: int,
                 ampl_v: float,
                 period_s: float,
                 width_s: float,
                 delay_s: float,
                 wait_count: int = 0,      # Nb de pulses à attendre
                 mode: str = "NORMAL",
                 polarity: str = "NORMAL",
                 gate: str = "DISABLE",
                 trigger: str = "DISABLED",) -> None:

    pulse_generator.write(f":PULSE{channel}:STATE ON")

    pulse_generator.write(f":PULSE{channel}:OUTP:AMPL {ampl_v:.6f}")
    pulse_generator.write(f":PULSE{channel}:WIDT {width_s:.12f}")
    pulse_generator.write(f":PULSE{channel}:DELAY {delay_s:.12f}")

    if wait_count > 0:
        pulse_generator.write(f":PULSE{channel}:WCOUNTER {wait_count}N")

    pulse_generator.write(f":PULSE{channel}:CMODE {mode}")
    pulse_generator.write(f":PULSE{channel}:POL {polarity}")
    pulse_generator.write(f":PULSE{channel}:CGATE {gate}")

    pulse_generator.write(f":PULSE0:EXT:MODE {trigger}")
    pulse_generator.write(":PULSE0:MODE NORM")
    pulse_generator.write(f":PULSE0:PER {period_s:.12f}")

    pulse_generator.write(":PULSE0:STATE ON")
    time.sleep(0.1)


class PulseThread(QThread):
    log_signal = Signal(str)
    finished_signal = Signal()
    error_signal = Signal(str)

    def __init__(self, channels, ampl_v, freq_mhz, width_s, delay_s, wait_count, mode, polarity, gate, trigger):
        super().__init__()
        self.channels    = channels
        self.ampl_v      = ampl_v
        self.freq_mhz    = freq_mhz
        self.width_s     = width_s
        self.delay_s     = delay_s
        self.wait_count  = wait_count
        self.mode        = mode
        self.polarity    = polarity
        self.gate        = gate
        self.trigger     = trigger

    def log(self, msg: str):
        self.log_signal.emit(msg)

    def run(self):
        rm = None
        pulse_generator = None
        try:
            rm, pulse_generator = Gpid_devices_open()
            Pulse_generator_init(pulse_generator)

            period_s = MHz_to_s(self.freq_mhz)

            for ch in [1, 2, 3, 4]:
                if self.channels.get(ch, False):
                    Create_pulse(
                        pulse_generator=pulse_generator,
                        channel=ch,
                        ampl_v=self.ampl_v,
                        period_s=period_s,
                        width_s=self.width_s,
                        delay_s=self.delay_s,
                        wait_count=self.wait_count,
                        mode=self.mode,
                        polarity=self.polarity,
                        gate=self.gate,
                        trigger=self.trigger,
                    )
                else:
                    pulse_generator.write(f":PULSE{ch}:STATE OFF")

            self.log("Done")
            self.finished_signal.emit()

        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            if pulse_generator is not None and rm is not None:
                CLOSE_ALL(pulse_generator, rm)


class StopAllThread(QThread):
    log_signal = Signal(str)
    finished_signal = Signal()
    error_signal = Signal(str)

    def log(self, msg: str):
        self.log_signal.emit(msg)

    def run(self):
        rm = None
        pulse_generator = None
        try:
            rm, pulse_generator = Gpid_devices_open()
            for ch in [1, 2, 3, 4]:
                pulse_generator.write(f":PULSE{ch}:STATE OFF")
            pulse_generator.write(":PULSE0:STATE OFF")
            self.log("STOP ALL: CH1..CH4 OFF + PULSE0 OFF")
            self.finished_signal.emit()
        except Exception as e:
            self.error_signal.emit(str(e))
        finally:
            if pulse_generator is not None and rm is not None:
                CLOSE_ALL(pulse_generator, rm)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Pulse Generator – BNC505")
        self.resize(600, 600)
        self.setMinimumSize(500, 450)

        base_dir = Path(__file__).resolve().parent
        self.logo_path = base_dir / "Logo_CIM.jpg"

        # Channels
        self.t1 = QCheckBox("T1"); self.t1.setChecked(True)
        self.t2 = QCheckBox("T2")
        self.t3 = QCheckBox("T3")
        self.t4 = QCheckBox("T4")

        ch_widget = QWidget()
        ch_layout = QHBoxLayout()
        ch_layout.setContentsMargins(0, 0, 0, 0)
        ch_layout.setSpacing(6)
        ch_layout.addWidget(self.t1)
        ch_layout.addWidget(self.t2)
        ch_layout.addWidget(self.t3)
        ch_layout.addWidget(self.t4)
        ch_layout.addStretch()
        ch_widget.setLayout(ch_layout)

        # ── Left column ──────────────────────────────────────────────
        self.ampl_edit = QLineEdit(); self.ampl_edit.setPlaceholderText("6.0")

        # Frequency
        self.freq_edit = QLineEdit(); self.freq_edit.setPlaceholderText("1.0")
        self.freq_unit_combo = QComboBox()
        self.freq_unit_combo.addItems(["Hz", "KHz", "MHz", "GHz"])
        self.freq_unit_combo.setCurrentText("MHz")
        self.freq_unit_combo.setFixedWidth(70)

        self.period_title = QLabel("Period:")
        self.period_value = QLineEdit("—")
        self.period_value.setReadOnly(True)
        self.period_value.setAlignment(Qt.AlignCenter)

        # Pulse width
        self.width_edit = QLineEdit(); self.width_edit.setPlaceholderText("200")
        self.width_unit_combo = QComboBox()
        self.width_unit_combo.addItems(["ns", "µs", "ms", "s"])
        self.width_unit_combo.setCurrentText("ns")
        self.width_unit_combo.setFixedWidth(70)

        # Delay
        self.delay_edit = QLineEdit(); self.delay_edit.setPlaceholderText("0.0")
        self.delay_unit_combo = QComboBox()
        self.delay_unit_combo.addItems(["ns", "µs", "ms", "s"])
        self.delay_unit_combo.setCurrentText("ns")
        self.delay_unit_combo.setFixedWidth(70)

        # Wait : entier (nombre de pulses)
        self.wait_edit = QLineEdit(); self.wait_edit.setPlaceholderText("0")

        # ── Right column ─────────────────────────────────────────────
        self.mode_combo = QComboBox(); self.mode_combo.addItems(["NORMAL", "SINGLE", "BURST", "DCYCLE"])
        self.pol_combo  = QComboBox(); self.pol_combo.addItems(["NORMAL", "COMPLEMENT"])
        self.gate_combo = QComboBox(); self.gate_combo.addItems(["DISABLE", "LOW", "HIGH"])
        self.trig_combo = QComboBox(); self.trig_combo.addItems(["DISABLED", "TRIGGER", "GATE"])

        # ── Helper : builds a (value QLineEdit + unit QComboBox) cell ─
        def make_cell(line_edit, combo) -> QWidget:
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(6)
            row.addWidget(line_edit)
            row.addWidget(combo)
            cell = QWidget()
            cell.setLayout(row)
            cell.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
            cell.setFixedWidth(180)
            return cell

        # ── Grid ─────────────────────────────────────────────────────
        grid = QGridLayout()
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(6)

        # Row 0 : Output
        grid.addWidget(QLabel("Output:"), 0, 0)
        grid.addWidget(ch_widget, 0, 1, 1, 3)

        # Row 1 : Amplitude | Mode
        grid.addWidget(QLabel("Amplitude Pulse (V):"), 1, 0)
        grid.addWidget(self.ampl_edit, 1, 1)
        grid.addWidget(QLabel("Mode:"), 1, 2)
        grid.addWidget(self.mode_combo, 1, 3)

        # Row 2 : Frequency | Period
        grid.addWidget(QLabel("Frequency:"), 2, 0)
        grid.addWidget(make_cell(self.freq_edit, self.freq_unit_combo), 2, 1, alignment=Qt.AlignVCenter)
        grid.addWidget(self.period_title, 2, 2, alignment=Qt.AlignVCenter)
        grid.addWidget(self.period_value, 2, 3)

        # Row 3 : Pulse width | Polarity
        grid.addWidget(QLabel("Pulse width:"), 3, 0)
        grid.addWidget(make_cell(self.width_edit, self.width_unit_combo), 3, 1, alignment=Qt.AlignVCenter)
        grid.addWidget(QLabel("Polarity:"), 3, 2)
        grid.addWidget(self.pol_combo, 3, 3)

        # Row 4 : Delay | Gate
        grid.addWidget(QLabel("Delay:"), 4, 0)
        grid.addWidget(make_cell(self.delay_edit, self.delay_unit_combo), 4, 1, alignment=Qt.AlignVCenter)
        grid.addWidget(QLabel("Gate:"), 4, 2)
        grid.addWidget(self.gate_combo, 4, 3)

        # Row 5 : Wait (Nb pulses) | Trigger
        grid.addWidget(QLabel("Wait (Nb pulses):"), 5, 0)
        grid.addWidget(self.wait_edit, 5, 1)
        grid.addWidget(QLabel("Trigger:"), 5, 2)
        grid.addWidget(self.trig_combo, 5, 3)

        # ── Logo ─────────────────────────────────────────────────────
        self.logo_label = QLabel()
        if self.logo_path.exists():
            pm = QPixmap(str(self.logo_path))
            pm = pm.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(pm)
        self.logo_label.setAlignment(Qt.AlignLeft)

        # ── Buttons ──────────────────────────────────────────────────
        self.run_button      = QPushButton("START")
        self.stop_all_button = QPushButton("STOP ALL")
        self.exit_button     = QPushButton("Quitter")

        self.run_button.clicked.connect(self.on_start)
        self.stop_all_button.clicked.connect(self.on_stop_all)
        self.exit_button.clicked.connect(self.close)

        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(self.stop_all_button)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.run_button)
        buttons_layout.addWidget(self.exit_button)

        # ── Log ──────────────────────────────────────────────────────
        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setStyleSheet("border: 1px solid #A0A0A0; font-family: 'Courier New', 'Consolas', monospace;")
        self.log_edit.setMaximumBlockCount(2000)
        self.log_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.log_edit.setMinimumHeight(150)

        # ── Main layout ──────────────────────────────────────────────
        layout = QVBoxLayout()
        layout.addLayout(grid)
        layout.addWidget(self.logo_label, 0, Qt.AlignLeft)
        layout.addLayout(buttons_layout)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        layout.addWidget(self.log_edit, 1)
        self.setLayout(layout)

        # Live period update
        self.freq_edit.textChanged.connect(self.update_period_label)
        self.freq_unit_combo.currentTextChanged.connect(self.update_period_label)
        self.update_period_label()

        self.thread = None
        self.run_id = 0

    # ------------------------------------------------------------------ helpers

    def log(self, message: str):
        self.log_edit.appendPlainText(message)
        self.log_edit.ensureCursorVisible()
        QCoreApplication.processEvents()

    def _channels_dict(self):
        return {1: self.t1.isChecked(), 2: self.t2.isChecked(),
                3: self.t3.isChecked(), 4: self.t4.isChecked()}

    def _freq_to_hz(self, value: float, unit: str) -> float:
        factors = {"Hz": 1, "KHz": 1e3, "MHz": 1e6, "GHz": 1e9}
        if unit not in factors:
            raise ValueError("Unknown frequency unit")
        return value * factors[unit]

    def _time_to_s(self, value: float, unit: str) -> float:
        factors = {"s": 1, "ms": 1e-3, "µs": 1e-6, "ns": 1e-9}
        if unit not in factors:
            raise ValueError("Unknown time unit")
        return value * factors[unit]

    def _format_time(self, seconds: float) -> str:
        if seconds >= 1:
            return f"{seconds:.6g} s"
        if seconds >= 1e-3:
            return f"{seconds * 1e3:.6g} ms"
        if seconds >= 1e-6:
            return f"{seconds * 1e6:.6g} µs"
        return f"{seconds * 1e9:.6g} ns"

    def update_period_label(self):
        text = (self.freq_edit.text() or "").strip()
        unit = self.freq_unit_combo.currentText()
        if not text:
            self.period_value.setText("—")
            return
        try:
            f_hz = self._freq_to_hz(float(text), unit)
            if f_hz <= 0:
                self.period_value.setText("—")
                return
            self.period_value.setText(self._format_time(1.0 / f_hz))
        except ValueError:
            self.period_value.setText("—")

    # ------------------------------------------------------------------ slots

    def on_start(self):
        channels = self._channels_dict()
        if not any(channels.values()):
            QMessageBox.warning(self, "Erreur", "Sélectionne au moins un canal (T1..T4).")
            return

        try:
            ampl_v = float(self.ampl_edit.text() or "6.0")

            freq_val  = float(self.freq_edit.text() or "1.0")
            freq_unit = self.freq_unit_combo.currentText()
            freq_hz   = self._freq_to_hz(freq_val, freq_unit)
            freq_mhz  = freq_hz / 1e6

            width_val  = float(self.width_edit.text() or "200")
            width_unit = self.width_unit_combo.currentText()
            width_s    = self._time_to_s(width_val, width_unit)

            delay_val  = float(self.delay_edit.text() or "0.0")
            delay_unit = self.delay_unit_combo.currentText()
            delay_s    = self._time_to_s(delay_val, delay_unit)

            wait_count = int(self.wait_edit.text() or "0")
            if wait_count < 0:
                raise ValueError("Wait (Nb pulses) doit être >= 0")

            mode     = self.mode_combo.currentText()
            polarity = self.pol_combo.currentText()
            gate     = self.gate_combo.currentText()
            trigger  = self.trig_combo.currentText()

        except ValueError as e:
            QMessageBox.warning(self, "Erreur saisie", f"Valeur invalide:\n{e}")
            return

        self.run_id += 1
        ts       = QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
        ch_mask  = " ".join([f"T{ch}" for ch in [1, 2, 3, 4] if channels[ch]])
        period_s = MHz_to_s(freq_mhz)

        K = 16
        V = 14

        self.log("")
        self.log("=" * 60)
        self.log(f"Run n°{self.run_id} : {ts}")
        self.log(f"Channels : {ch_mask}")
        self.log("-" * 60)
        self.log(f"{'Ampl (V)':<{K}}: {ampl_v:<{V}.{PRECISION}f}  {'Mode':<{K}}: {mode}")
        self.log(f"{'Freq':<{K}}: {f'{freq_val} {freq_unit}':<{V}}  {'Period':<{K}}: {self._format_time(period_s)}")
        self.log(f"{'Width':<{K}}: {f'{width_val} {width_unit}':<{V}}  {'Polarity':<{K}}: {polarity}")
        self.log(f"{'Delay':<{K}}: {f'{delay_val} {delay_unit}':<{V}}  {'Gate':<{K}}: {gate}")
        self.log(f"{'Wait (Nb pulses)':<{K}}: {wait_count:<{V}}  {'Trigger':<{K}}: {trigger}")
        self.log("=" * 60)

        self.run_button.setEnabled(False)
        self.stop_all_button.setEnabled(False)

        self.thread = PulseThread(
            channels=channels,
            ampl_v=ampl_v,
            freq_mhz=freq_mhz,
            width_s=width_s,
            delay_s=delay_s,
            wait_count=wait_count,
            mode=mode,
            polarity=polarity,
            gate=gate,
            trigger=trigger,
        )
        self.thread.log_signal.connect(self.log)
        self.thread.finished_signal.connect(self.on_thread_finished)
        self.thread.error_signal.connect(self.on_thread_error)
        self.thread.start()

    def on_stop_all(self):
        self.run_button.setEnabled(False)
        self.stop_all_button.setEnabled(False)

        self.thread = StopAllThread()
        self.thread.log_signal.connect(self.log)
        self.thread.finished_signal.connect(self.on_thread_finished)
        self.thread.error_signal.connect(self.on_thread_error)
        self.thread.start()

    def on_thread_finished(self):
        self.log("✅ OK")
        self.run_button.setEnabled(True)
        self.stop_all_button.setEnabled(True)

    def on_thread_error(self, msg: str):
        self.log(f"❌ ERREUR: {msg}")
        self.run_button.setEnabled(True)
        self.stop_all_button.setEnabled(True)
        QMessageBox.critical(self, "Erreur", msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    base_dir = Path(__file__).resolve().parent

    app_icon_path = base_dir / "BNC-505.png"
    if app_icon_path.exists():
        app.setWindowIcon(QIcon(str(app_icon_path)))

    win = MainWindow()
    win.show()
    sys.exit(app.exec())

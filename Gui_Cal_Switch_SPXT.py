# By Arthur Péraud 12/2025
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget, QFormLayout, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QMessageBox, QPlainTextEdit
)
from PySide6.QtCore import Qt, QCoreApplication
from PySide6.QtGui import QIcon,QPixmap
from PySide6.QtWidgets import QLabel

from Cal_Switch_SPXT import (
    Ltoi,
    Find_sheet_index,
    Read_mdb_table,
    Get_col_from_mdb,
    Get_unique_filename,
    Add_offset,
    Fill_sheet_from_channel,
    Fill_voies_sheets,
    Build_path_names
)


def Save_workbook_gui(parent_widget, wb, output_file: str) -> bool:
    """
    Retourne True si un fichier a été sauvegardé, False sinon.
    """
    output_file = str(output_file)
    path = Path(output_file)

    if path.exists():
        reply = QMessageBox.question(
            parent_widget,
            "Fichier existant",
            f"Le fichier\n\n{output_file}\n\nexiste déjà.\n"
            "Voulez-vous l'écraser ?",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Cancel
        )

        if reply == QMessageBox.Yes:
            wb.save(output_file)
            QMessageBox.information(
                parent_widget, "Sauvegarde effectuée",
                f"Fichier écrasé et sauvegardé sous :\n{output_file}"
            )
            return True

        elif reply == QMessageBox.No:
            new_file = Get_unique_filename(output_file)
            wb.save(new_file)
            QMessageBox.information(
                parent_widget, "Sauvegarde effectuée",
                f"Fichier sauvegardé sous un nouveau nom :\n{new_file}"
            )
            return True

        else:
            # Quitter OU croix
            return False

    else:
        wb.save(output_file)
        QMessageBox.information(
            parent_widget, "Sauvegarde effectuée",
            f"Fichier sauvegardé sous :\n{output_file}"
        )
        return True


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Cal Info Mesure – SWITCH SPXT")

        # Logo
        # Suppose que Logo_CIM.jpg est dans le même dossier que ce script
        base_dir = Path(__file__).resolve().parent
        icon_path = base_dir / "Logo_CIM.jpg"
        self.setWindowIcon(QIcon(str(icon_path)))

        self.year_edit = QLineEdit()
        self.year_edit.setPlaceholderText("Entrez l'année (ex : 2025)")

        self.client_edit = QLineEdit()
        self.client_edit.setPlaceholderText("Entrez le nom du client")

        self.freqband_edit = QLineEdit()
        self.freqband_edit.setPlaceholderText("Entrez la bande (XXXX)")
        self.freqband_edit.setMaxLength(4)

        self.sn_edit = QLineEdit()
        self.sn_edit.setPlaceholderText("Entrez le n° de série (XXXX)")
        self.sn_edit.setMaxLength(4)

        form = QFormLayout()
        form.addRow("Année :", self.year_edit)
        form.addRow("Client :", self.client_edit)
        form.addRow("Bande de fréquence :", self.freqband_edit)
        form.addRow("N° de série :", self.sn_edit)

        self.run_button = QPushButton("OK")
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

        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(logo_label)    
        buttons_layout.addWidget(self.run_button)
        buttons_layout.addWidget(self.exit_button)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)

        layout = QVBoxLayout()
        layout.addLayout(form)
        layout.addLayout(buttons_layout)
        layout.addWidget(self.log_edit)

        self.setLayout(layout)
        self.resize(500, 300)

    def log(self, message: str):
        self.log_edit.appendPlainText(message)
        QCoreApplication.processEvents()

    def on_ok_clicked(self):
        year_text = self.year_edit.text().strip()
        client = self.client_edit.text().strip()
        freqband = self.freqband_edit.text().strip()
        sn = self.sn_edit.text().strip()

        if not year_text.isdigit():
            QMessageBox.warning(self, "Erreur", "L'année doit être un entier.")
            return
        if len(freqband) != 4 or not freqband.isdigit():
            QMessageBox.warning(self, "Erreur", "La bande doit être 4 chiffres (XXXX).")
            return
        if len(sn) != 4 or not sn.isdigit():
            QMessageBox.warning(self, "Erreur", "Le numéro de série doit être 4 chiffres (XXXX).")
            return

        year = int(year_text)

        try:
            # MAIN
            data_path1, data_path2, output_file, input_file = Build_path_names(
                client, year, freqband, sn
            )

            wb = Fill_voies_sheets(input_file, data_path1, data_path2, log_func=self.log)
            saved = Save_workbook_gui(self, wb, str(output_file))

            if saved:
                QMessageBox.information(
                    self,
                    "Succès",
                    f"Traitement terminé.\nFichier généré :\n{output_file}"
                )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Erreur",
                f"Une erreur s'est produite :\n{e}"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    base_dir = Path(__file__).resolve().parent
    icon_path = base_dir / "Logo_CIM.png"
    app.setWindowIcon(QIcon(str(icon_path)))
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


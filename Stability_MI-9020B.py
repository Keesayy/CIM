import os, openpyxl, sys, subprocess, csv, io 
import pandas as pd
import numpy as np
import time
import pyvisa

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pathlib import Path

PRECISION = 3
NB_DIGIT = 4     # Power Meter
NB_ESE_BITS = 60
NB_QUES_BITS = 952
DEBUG = False

def PRINT(*args, **kwargs):
	if(DEBUG):
		return __builtins__.print(*args, **kwargs)

def Hz_to_GHz(x : float) -> float:
	new = x * 1E9
	return new

def Float_precision_str(n : int) -> str:
	s : str = '#,##0.'
	for i in range(n):
		s += '0'
	return s

def Excel_name(name : str, precision : int, freq : float, amp : float) -> str:
	def format_amp(amp: float) -> str:
		if amp < 0:
		    return f"Neg{int(abs(amp))}"
		return str(int(amp))

	s = name
	s += f"_{format_amp(freq)}"
	s += f"_{format_amp(amp)}"
	s += f'_MI-9020B.xlsx'
	return s

def Get_unique_filename(path : str) -> str:
    """Génère un nom de fichier unique en ajoutant (1), (2), etc. s'il existe déjà."""
    base, ext = os.path.splitext(path)
    i = 1
    new_path = path
    while os.path.exists(new_path):
        new_path = f"{base}({i}){ext}"
        i += 1
    return new_path

def Build_path_names(name : str, client: str, year: int) -> tuple[Path, Path, Path, Path]:
    base_dir = Path(rf"E:\\Cal Info Mesure\\{client}\\Data {year}")
    output_file = base_dir / name
    return output_file

def Save_workbook_safely(wb : Workbook, output_file: str) -> None:
    """Sauvegarde fichier excel, confirmation si existence sinon préfixe est ajouté (Get_unique_filename)."""
    if os.path.exists(output_file):
        print(f"⚠️  Le fichier '{output_file}' existé déjà")
        new_file = Get_unique_filename(output_file)
        wb.save(new_file)
        print(f"📁 Fichier sauvegardé sous un nouveau nom : {new_file}")
    else:
        wb.save(output_file)
        print(f"✅ Fichier sauvegardé sous {output_file}")

def Gpid_devices_open():
	rm = pyvisa.ResourceManager()
	print(rm.list_resources(), '\n')

	#Open gpid devices
	power_meter = rm.open_resource('GPIB0::13::INSTR')
	signal_source = rm.open_resource('TCPIP::192.168.10.100::INSTR')

	power_meter.write('SYST:LANG SCPI')
	signal_source.write('SYST:LANG SCPI')
	time.sleep(0.2)

	print(power_meter.query('*IDN?'), end = "")
	print(signal_source.write('*IDN?'), end = "")
	return rm ,power_meter, signal_source

### Signal Source Init
def Signal_source_init(signal_source) -> None:
	# signal_source.write('*RST')
	signal_source.write('*CLS')
	signal_source.write('*ESE 0')
	signal_source.write('*SRE 0')

### Power Meter Init
def Power_meter_init(power_meter) -> None:
	power_meter.write('*CLS')
	power_meter.write('*ESE 1') 
	power_meter.write('UNIT:POW dBm')

	PRINT('ESR : ', power_meter.write('*ESR?'))

	power_meter.write('DISP:RES %d' % NB_DIGIT)

def Show_parameters(freq : float, dwel : float, amplitude : float, t : float) -> None:
	print('\nSTARTING ACQUISITION WITH PARAMETERS :')
	print('freq :', ('{:.%df}' % PRECISION).format(freq), 'Ghz')
	print('dwel   :', ('{:.%df}' % PRECISION).format(dwel), 'ms')
	print('amp    :', ('{:.%df}' % PRECISION).format(amplitude), 'dBm')
	print('temps    :', ('{:.%df}' % PRECISION).format(t), 's')
	print('\n')


def Stability_test(power_meter, signal_source, freq : float, dwel : float, amp : float, t_tot : float, t : float) -> Workbook:
	freq = Hz_to_GHz(freq)

	Show_parameters(freq, dwel, amp, t_tot)
	
	excel = openpyxl.Workbook()
	sheet = excel.active
	sheet.title = "Data"
	sheet['B' + str(1)] = 'Fréquence (GHz)'
	sheet['C' + str(1)] = f'{amp} dBm'

	# signal_source.write('OUTP ON')
	PRINT(signal_source.query('*OPC?'))
	signal_source.write('POW %f dBm' % amp)

	precision_string = Float_precision_str(PRECISION)

	start_total = time.time()
	t_0 = 0

	measure_count = 0
	while t_0 < t_tot:
	    signal_source.write('FREQ:CW %f' % freq)

	    power_meter.write('*CLS')
	    power_meter.write('FREQ ' + str(freq))

	    #Measure level
	    power_meter.write('TRIG:DEL:AUTO ON')
	    power_meter.write('INIT:CONT OFF') 
	    power_meter.write('TRIG:SOUR IMM') 
	    power_meter.write('INIT') 

	    power_meter.write('*OPC')

	    time.sleep(t)

	    #Clear ESE
	    power_meter.query('*ESR?') 

	    #Read Level
	    level = power_meter.query('FETCH?')
	    
	    # TEMPS ÉCOULÉ EN SECONDES
	    elapsed = time.time() - start_total
	    t_0 += t

	    measure_count += 1
	    print(f"{measure_count:4d}: {elapsed:7.0f}s | "
	          f"{float(freq):.3f} GHz | {float(level):.3f} dBm")mm

	    #Excel array
	    row = measure_count + 2
	    sheet['B' + str(row)] = t_0  # Secondes totales
	    sheet['C' + str(row)] = float(level)
	    sheet['B' + str(row)].number_format = precision_string
	    sheet['C' + str(row)].number_format = precision_string

	print(f"FIN {measure_count} mesures en {elapsed:.0f}s")

	return excel

def CLOSE_ALL(signal_source, power_meter, excel, rm) -> None:
	signal_source.close()
	power_meter.close()
	excel.close()
	rm.close()	

def main() -> int:
	print("Test Stabilité Power Meter _MI-9020B.")
	rm, power_meter, signal_source = Gpid_devices_open()

	Signal_source_init(signal_source)
	Power_meter_init(power_meter)

	client = input("Entrez nom du client : ")
	year = int(input("Entrez l'année : "))
	freq = float(input("Entrez la fréquence en GHZ : "))
	amp = float(input("Entrez l'amplitude en dBm : "))
	t_tot = float(input("Entrez le temps total en s : "))
	t = float(input("Entrez le pas de temps en s : "))

	# Step_Sweep
	# freq     : float = 0.01	#GHZ	
	dwel     : float = 1.00	#MS
	# amp      : float = -30  #DBM
	# t_tot 	: float = 3600 #s
	# t 	     : float = 5 #s

	excel_name = Excel_name('Stability', 0, freq, amp)	
	excel = Stability_test(power_meter, signal_source, freq, dwel, amp, t_tot, t)

	output_file = Build_path_names(excel_name, client, year)
	Save_workbook_safely(excel, output_file)

	CLOSE_ALL(signal_source, power_meter, excel, rm)
	return 0

if __name__=="__main__":
    main()








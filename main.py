# Границі основної приведеної похибки - звести розмір до 2 ячейок у всіх протоколах
# Установку дефолтних значений и настроек перенесы в методы работы
# висоту заголовків таблиці повірки 45 пк всюди
# Для MTE 143 нема 4 проводної поправ таблицю

# https://www.ibm.com/developerworks/ru/library/l-python_part_6/index.html
# https://wiki.qt.io/PySideSimplicissimus_Module_2_CloseButton
# https://www.blog.pythonlibrary.org/2018/05/30/loading-ui-files-in-qt-for-python/
# https://wiki.qt.io/Qt_for_Python_UiFiles

import openpyxl
import sys
import time
#import threading

import config

from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QPushButton, QLineEdit, QComboBox, QTextBrowser
from PySide2.QtCore import QFile, QObject



class main(QObject):
	#-----------
	# Подключение GUI
	#----------- 	
	def __init__(self, ui_file, parent=None):
		super(main, self).__init__(parent)
		ui_file = QFile(ui_file)
		ui_file.open(QFile.ReadOnly)

		loader = QUiLoader()
		self.window = loader.load(ui_file)
		ui_file.close()

		#-----------
		# Элементы формы
		#-----------
	
		# Окно вывода
		self.outBox = self.window.findChild(QTextBrowser, 'outBox')

		# Настройки программы
		self.comNumber = self.window.findChild(QLineEdit, 'comNumber')
		
		# Данные прибора
		self.targetType = self.window.findChild(QComboBox, 'targetType')
		self.targetNumber = self.window.findChild(QLineEdit, 'targetNumber')		
		self.targetVar = self.window.findChild(QLineEdit, 'targetVar')
		self.targetExNumber = self.window.findChild(QComboBox, 'targetExNumber')
		self.targetExParam = self.window.findChild(QComboBox, 'targetExParam')
		self.targetExCode = self.window.findChild(QComboBox, 'targetExCode')
		
		# Настройки поверки
		self.protNumber = self.window.findChild(QLineEdit, 'protNumber')
		self.warmupTime = self.window.findChild(QLineEdit, 'warmupTime')
		self.oneStepTime = self.window.findChild(QLineEdit, 'oneStepTime')
		
		# Управление
		str_btn = self.window.findChild(QPushButton, 'startButton')
		stp_btn = self.window.findChild(QPushButton, 'stopButton')
		
		# События формы
		str_btn.clicked.connect(self.start_button)
		stp_btn.clicked.connect(self.stop_button)		
		
		self.window.show()
		
		

	#-----------
	# Обработка событий формы
	#-----------
	def start_button(self):
		self.outBox.clear()
		print('Перевод в режим дистанционного управления')
		config.ser.write(b'SYST:REM \n')
		device_start_execute = getattr(self, self.targetType.currentText().replace("/","_")+'_'+self.targetExParam.currentText())
		device_start_execute()
		self.stop_button()
		
		#global my_thread, my_thread_stop		
		#my_thread_stop = False
		#my_thread = threading.Thread(target=self.sex)
		#my_thread.start()
		
	#def sex(self):
	#	for _ in range(10):
	#		time.sleep(1)
	#		print(threading.currentThread().getName() + '\n')
	#		print(my_thread_stop)
	#		my_thread.join()
		
		
	def stop_button(self):
		#print('Сброс сигнала')
		#global my_thread_stop
		#my_thread_stop = True
		config.ser.write(b"OUTP OFF \n")
		config.ser.write(b"SYST:LOC \n")
	
	#-----------
	# Методы
	#-----------
	def prot(self, case, sheet_col_init, sheet_row_init):
		# Подготовка протокола
		if case=='init':
			global wb
			global sheet
			global sheet_col
			global sheet_row

			sheet_col = sheet_col_init
			sheet_row = sheet_row_init
			wb = openpyxl.load_workbook(filename="Libary/"+self.targetType.currentText().replace("/","#")+'_'+self.targetExParam.currentText()+".xlsx")
			sheet = wb.worksheets[0]
			
			# Ввод данных протокола
			sheet['L1'] = self.protNumber.text().replace("#","/")
			sheet['R3'] = self.targetNumber.text()
			sheet['B20'] = self.targetExNumber.currentText()
			sheet['E20'] = self.targetExParam.currentText()
			sheet['H20'] = self.targetExCode.currentText()
			sheet['I5'] = self.targetVar.text().replace(".",",")
			sheet['T63'] = time.strftime("%d.%m.%Y", time.gmtime())
			
		# Ввод данных поверки
		elif case=='data':
			config.ser.write(b'MEAS? \n')			
			OUT = config.ser.read(17).strip().decode("utf-8")
			# Преобразование выходных данных
			OUT_point = OUT.find('e')
			EXP = OUT[OUT_point+1:]
			EXP = int(EXP)+3
			OUT = float(OUT[:6])*10**EXP
			OUT = str(OUT)
			OUT = OUT[:6].replace(".",",")
			
			# Занос в протокол
			print(OUT)
			sheet[sheet_col+str(sheet_row)] = OUT
			sheet_row = int(sheet_row)
			sheet_row += 1		
		
		# Вывод протокола
		elif case =='end':
			wb.save(filename="Protocols/"+self.protNumber.text()+".xlsx")
			print("Конец поверки")			
	
	def warmup(self, mode_warmup, time_warmup, perc_warmup, nom_var):
		print("Прогрев")
		# Подготовка данных
		mode_warmup = str(mode_warmup).encode('ascii')
		var_warmup = str(nom_var*perc_warmup).encode('ascii')
			
		# Прогрев
		config.ser.write(mode_warmup+b" "+var_warmup+b" \n")
		config.ser.write(b"OUTP ON \n")
		time.sleep(time_warmup)
		print("Поверка")

	#-----------
	# Метод работы
	#-----------
	def AC_current(self, time_warmup, perc_warmup, time_step, I_nom, I_steps):
	
		# Прогрев
		self.warmup("CAC:CURR", time_warmup, perc_warmup, I_nom)
	
		# Поверка
		i = 0
		while( i < len(I_steps)):
			config.ser.write(b'CAC:CURR '+(str(I_steps[i]).encode('ascii'))+b' \n')
			
			time.sleep(time_step)
			self.prot('data',0,0)
			time.sleep(2)
			
			i += 1
			
		self.prot('end',0,0)

	def AC_voltage(self, case, time_warmup, perc_warmup, time_step, U_nom, U_steps, F_nom, F_steps):
		
		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"VAC:VOLT "+(str(U_nom).encode('ascii'))+b" \n")
		config.ser.write(b"VAC:FREQ "+(str(F_nom).encode('ascii'))+b" \n")	
	
		# Прогрев
		if case=='volt':
			self.warmup("VAC:VOLT", time_warmup, perc_warmup, U_nom)
		
		if case=='freq':
			self.warmup("VAC:FREQ", time_warmup, perc_warmup, F_nom)
			
		# Поверка
		i = 0
		while( i < len(U_steps)):
			config.ser.write(b'VAC:VOLT '+(str(U_steps[i]).encode('ascii'))+b' \n')
			config.ser.write(b'VAC:FREQ '+(str(F_steps[i]).encode('ascii'))+b' \n')
			
			# Перезапуск подачи при переходе от значений меньших за 100 V к большим
			if ((U_steps[i] >= 100) and (U_steps[i-1] < U_steps[i])):
				config.ser.write(b"OUTP ON \n")
						
			time.sleep(time_step)
			self.prot('data',0,0)
			time.sleep(2)
			
			i += 1
			
		self.prot('end',0,0)
		
	def Power(self, time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps):
	
		# Прогрев
		self.warmup("PAC:CURR", time_warmup, perc_warmup, I_nom)

		# Поверка
		i = 0
		while( i < len(I_steps)):
			config.ser.write(b'PAC:CURR '+(str(I_steps[i]).encode('ascii'))+b' \n')
			config.ser.write(b'PAC:VOLT '+(str(U_steps[i]).encode('ascii'))+b' \n')
			config.ser.write(b'PAC:PHAS '+(str(D_steps[i]).encode('ascii'))+b' \n')
			
			time.sleep(time_step)
			self.prot('data',0,0)
			time.sleep(2)
			
			i += 1
			
		self.prot('end',0,0)
		
	#-----------
	# Приборы
	#----------

	def Test_F(self):
		print("OK")
		#wb = openpyxl.load_workbook(filename="Libary/Test.xlsx")
		#sheet = wb.worksheets[0]
		#sheet.oddHeader.left.text = "123"
		#wb.save(filename="Protocols/Test.xlsx")
		

		
		

	def E842_I(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		perc_warmup = 0.8
		
		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"CAC:FREQ 50 \n")
		
		# Шаги
		I_steps = []
		I_steps.append (I_nom * 0.2)
		I_steps.append (I_nom * 0.4)
		I_steps.append (I_nom * 0.6)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 1.0)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 0.6)
		I_steps.append (I_nom * 0.4)
		I_steps.append (I_nom * 0.2)
		
		self.prot('init','I','23')
		self.AC_current(time_warmup, perc_warmup, time_step, I_nom, I_steps)		
		
	def MTE_121_I(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"CAC:FREQ 50 \n")
		
		# Шаги
		I_steps = []
		I_steps.append (I_nom * 0.05)
		I_steps.append (I_nom * 0.2)
		I_steps.append (I_nom * 0.5)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 1.0)
		I_steps.append (I_nom * 1.1)

		self.prot('init','I','23')		
		self.AC_current(time_warmup, perc_warmup, time_step, I_nom, I_steps)
		
	def MTE_143_P(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"PAC:UNIT W \n")
		config.ser.write(b"PAC:VOLT 57.735 \n")
		config.ser.write(b"PAC:FREQ 50 \n")
		config.ser.write(b"PAC:PHAS 0 \n")
		config.ser.write(b"OUTP:CONF 123 \n")
		
		#Шаги
		I_steps = [];
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 0.01)
		I_steps.append (I_nom * 0.1)
		I_steps.append (I_nom * 0.5)
		I_steps.append (I_nom * 1.1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)		

		U_steps = [];
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.5)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1.1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		
		D_steps = [];
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (240)
		D_steps.append (330)
		D_steps.append (60)
		D_steps.append (150)		
		
		self.prot('init','P','23')
		self.Power(time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps)			

	def MTE_143_Q(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"PAC:UNIT VAR \n")
		config.ser.write(b"PAC:VOLT 57.735 \n")
		config.ser.write(b"PAC:FREQ 50 \n")
		config.ser.write(b"PAC:PHAS 0 \n")
		config.ser.write(b"OUTP:CONF 123 \n")
		
		#Шаги
		I_steps = [];
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 0.01)
		I_steps.append (I_nom * 0.1)
		I_steps.append (I_nom * 0.5)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)		

		U_steps = [];
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.5)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		
		D_steps = [];
		D_steps.append (0)
		D_steps.append (270)
		D_steps.append (90)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (240)
		D_steps.append (330)
		D_steps.append (300)
		D_steps.append (150)
		
		self.prot('init','P','23')		
		self.Power(time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps)
		
	def MTE_143_I(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"PAC:UNIT W \n")
		config.ser.write(b"PAC:VOLT 57.735 \n")
		config.ser.write(b"PAC:FREQ 50 \n")
		config.ser.write(b"PAC:PHAS 0 \n")
		config.ser.write(b"OUTP:CONF 123 \n")
		
		#Шаги
		I_steps = [];
		I_steps.append (I_nom * 0.05)
		I_steps.append (I_nom * 0.2)
		I_steps.append (I_nom * 0.5)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1.1)

		U_steps = [];
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		
		D_steps = [0, 0, 0, 0, 0, 0];

		self.prot('init','I','23')	
		self.Power(time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps)

	def MTE_142_P(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"PAC:UNIT W \n")
		config.ser.write(b"PAC:VOLT 57.735 \n")
		config.ser.write(b"PAC:FREQ 50 \n")
		config.ser.write(b"PAC:PHAS 0 \n")
		config.ser.write(b"OUTP:CONF 123 \n")
		
		#Шаги
		I_steps = [];
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 0.01)
		I_steps.append (I_nom * 0.1)
		I_steps.append (I_nom * 0.5)
		I_steps.append (I_nom * 1.1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)		

		U_steps = [];
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.5)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1.1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		
		D_steps = [];
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (0)
		D_steps.append (240)
		D_steps.append (330)
		D_steps.append (60)
		D_steps.append (150)	
		
		self.prot('init','P','23')
		self.Power(time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps)

	def MTE_142_Q(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"PAC:UNIT VAR \n")
		config.ser.write(b"PAC:VOLT 57.735 \n")
		config.ser.write(b"PAC:FREQ 50 \n")
		config.ser.write(b"PAC:PHAS 0 \n")
		config.ser.write(b"OUTP:CONF 123 \n")
		
		#Шаги
		I_steps = [];
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 0.01)
		I_steps.append (I_nom * 0.1)
		I_steps.append (I_nom * 0.5)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1)

		U_steps = [];
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.5)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		
		D_steps = [];
		D_steps.append (0)
		D_steps.append (270)
		D_steps.append (90)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (270)
		D_steps.append (240)
		D_steps.append (330)
		D_steps.append (300)
		D_steps.append (150)
		
		self.prot('init','P','23')
		self.Power(time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps)				
		
	def MTE_142_I(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"PAC:UNIT W \n")
		config.ser.write(b"PAC:VOLT 57.735 \n")
		config.ser.write(b"PAC:FREQ 50 \n")
		config.ser.write(b"PAC:PHAS 0 \n")
		config.ser.write(b"OUTP:CONF 123 \n")
		
		#Шаги
		I_steps = [];
		I_steps.append (I_nom * 0.05)
		I_steps.append (I_nom * 0.2)
		I_steps.append (I_nom * 0.5)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 1.1)

		U_steps = [];
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1)
		
		D_steps = [0, 0, 0, 0, 0, 0];	
		
		self.prot('init','I','23')
		self.Power(time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps)	
	
	def MTE_111_U(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		U_nom = int(self.targetVar.text())
		F_nom = 50
		perc_warmup = 0.8
	
		#Шаги
		U_steps = [];
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.5)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1.1)
		F_steps = [F_nom, F_nom, F_nom, F_nom, F_nom];

		self.prot('init','H','23')		
		self.AC_voltage('volt', time_warmup, perc_warmup, time_step, U_nom, U_steps, F_nom, F_steps)	
		
	def MTE_111_F(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		U_nom = int(self.targetVar.text())
		F_nom = 50
		perc_warmup = 1.06
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [45, 47.5, 50, 52.5, 55];

		self.prot('init','E','22')		
		self.AC_voltage('freq', time_warmup, perc_warmup, time_step, U_nom, U_steps, F_nom, F_steps)	

	def E858_1_F(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		U_nom = int(self.targetVar.text())
		F_nom = 50
		perc_warmup = 1
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [45, 45.01, 47.5, 50, 52.5, 54.99, 55];

		self.prot('init','F','22')		
		self.AC_voltage('freq', time_warmup, perc_warmup, time_step, U_nom, U_steps, F_nom, F_steps)		
		
	def E858_2_F(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		U_nom = int(self.targetVar.text())
		F_nom = 50
		perc_warmup = 1
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [48, 49, 50, 51, 52];

		self.prot('init','F','22')		
		self.AC_voltage('freq', time_warmup, perc_warmup, time_step, U_nom, U_steps, F_nom, F_steps)			
		
	def E855_1_U(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		U_nom = int(self.targetVar.text())
		F_nom = 50
		perc_warmup = 0.8
	
		#Шаги
		U_steps = [];
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.4)
		U_steps.append (U_nom * 0.6)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		F_steps = [F_nom, F_nom, F_nom, F_nom, F_nom];

		self.prot('init','H','23')		
		self.AC_voltage('volt', time_warmup, perc_warmup, time_step, U_nom, U_steps, F_nom, F_steps)

	def E855_2_U(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		U_nom = int(self.targetVar.text())
		F_nom = 50
		perc_warmup = 0.92
	
		#Шаги
		U_steps = [75, 85, 95, 105, 115, 125];
		F_steps = [F_nom, F_nom, F_nom, F_nom, F_nom, F_nom];

		self.prot('init','H','23')		
		self.AC_voltage('volt', time_warmup, perc_warmup, time_step, U_nom, U_steps, F_nom, F_steps)			
	
	def E848_P(self):
		# Вводные
		time_warmup = int(self.warmupTime.text())
		time_step = int(self.oneStepTime.text())
		I_nom = float(self.targetVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"PAC:UNIT W \n")
		config.ser.write(b"PAC:VOLT 57.735 \n")
		config.ser.write(b"PAC:FREQ 50 \n")
		config.ser.write(b"PAC:PHAS 0 \n")
		config.ser.write(b"OUTP:CONF 123 \n")
		
		#Шаги
		I_steps = [];
		I_steps.append (I_nom * 0.2)
		I_steps.append (I_nom * 0.4)
		I_steps.append (I_nom * 0.6)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 0.2)
		I_steps.append (I_nom * 0.4)
		I_steps.append (I_nom * 0.6)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 1)
		I_steps.append (I_nom * 0.2)
		I_steps.append (I_nom * 0.4)
		I_steps.append (I_nom * 0.6)
		I_steps.append (I_nom * 0.8)
		I_steps.append (I_nom * 1)		

		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom];
		D_steps = [0, 0, 0, 0, 0, 60, 60, 60, 60, 60, 300, 300, 300, 300, 300];
		
		self.prot('init','M','23')
		self.Power(time_warmup, perc_warmup, time_step, I_nom, I_steps, U_steps, D_steps)		

	E859_P = E848_P	
		
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = main('gui.ui')
    sys.exit(app.exec_())
# Author - ExOmRa
# License - CC BY-NC-ND 4.0 - https://creativecommons.org/licenses/by-nc-nd/4.0/
# Email - ExOmRa@gmail.com
# Repository - https://github.com/ExOmRa/

import openpyxl
import sys
import time
import math

import config

from concurrent.futures import ThreadPoolExecutor

from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QPushButton, QLineEdit, QComboBox, QTextBrowser, QButtonGroup
from PySide2.QtCore import QFile, QObject



class main(QObject):
	#-----------
	# Основные переменные
	#----------- 			
	executor = ThreadPoolExecutor(max_workers = 2)
	statusCode = 'init'

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
		self.targetType_Main = self.window.findChild(QComboBox, 'targetType_Main')
		self.targetType_Sub = self.window.findChild(QLineEdit, 'targetType_Sub')
		self.targetNumber = self.window.findChild(QLineEdit, 'targetNumber')		
		self.targetNomVar = self.window.findChild(QLineEdit, 'targetNomVar')
		self.targetExitNumber = self.window.findChild(QButtonGroup, 'targetExitNumber')
		self.targetExitParam = self.window.findChild(QComboBox, 'targetExitParam')
		self.targetExitCode = self.window.findChild(QComboBox, 'targetExitCode')
		self.targetAccuracy_Link = self.window.findChild(QLineEdit, 'targetAccuracy')
		
		
		# Настройки поверки
		self.protNumber_1 = self.window.findChild(QLineEdit, 'protNumber_1')
		self.protNumber_2 = self.window.findChild(QLineEdit, 'protNumber_2')
		self.time_warmup_Link = self.window.findChild(QLineEdit, 'warmupTime')
		self.time_step_Link = self.window.findChild(QLineEdit, 'oneStepTime')
		self.measureTime = self.window.findChild(QLineEdit, 'measureTime')

		# Управление
		self.window.findChild(QPushButton, 'startButton').clicked.connect(self.start_button)
		self.window.findChild(QPushButton, 'stopButton').clicked.connect(self.stop_button)
		self.window.findChild(QPushButton, 'pauseButton').clicked.connect(self.pause_button)
		self.window.findChild(QPushButton, 'testButton').clicked.connect(self.test_button)
		
		# События формы
		self.window.show()
		
		

	#-----------
	# Обработка событий формы
	#-----------
	def start_button(self):
		self.outBox.clear()
		self.outBox.append('Перехід в режим дистанційного управління')
		config.ser.write(b'SYST:REM \n')
		
		# Общие параметры
		self.targetAccuracy = float(self.targetAccuracy_Link.text().replace(",","."))
		self.time_warmup = int(self.time_warmup_Link.text())
		self.time_step = int(self.time_step_Link.text())

		device_start_execute = getattr(self, self.targetType_Main.currentText().replace("/","_")+'_'+self.targetExitParam.currentText())
		if (self.targetExitCode.currentText() == '0..20'):
			self.targetExitValue_Nom = 20
		else:
			self.targetExitValue_Nom = 5

		# Инициализация
		if self.statusCode == 'init' or self.statusCode == 'stop':
			self.statusCode = 'working'
			self.executor.submit(device_start_execute)

	def stop_button(self):
		self.stop()
		self.outBox.append('Сброс сигнала')

	def pause_button(self):
		self.statusCode = 'pause'
		#print(self.targetExitNumber.checkedButton().text())


	def test_button(self):
		self.outBox.clear()
		self.outBox.append('Перехід в режим дистанційного управління')
		config.ser.write(b'SYST:REM \n')

		config.ser.write(b'MEAS? \n')
		OUT = config.ser.read(17).strip().decode("utf-8")
		print(OUT)
		# Преобразование выходных данных
		OUT_point = OUT.find('e')
		EXP = OUT[OUT_point+1:]
		EXP = int(EXP)+3
		OUT = float(OUT[:6])*10**EXP
		OUT = str(OUT)
		OUT = OUT[:6].replace(".",",")
		print(OUT)

	#-----------
	# Методы
	#-----------
	def stop(self):
		config.ser.write(b"OUTP OFF \n")
		config.ser.write(b"SYST:LOC \n")
		self.statusCode = 'stop'		

	def flagSleep(self, timeToSleep):
		for i in range(timeToSleep):
			if self.statusCode == 'stop':
				sys.exit()
			time.sleep(1)

	def accuracyCalc(self, OUT_series):
			# Поиск наибольшего значения
			OUT_series.sort()
			OUT_max = OUT_series[len(OUT_series) - 1]
			OUT_min = OUT_series[0]

			# Подсчет погрешности
			if (self.targetExitParam.currentText() == 'F'):
				calc_Max = ((OUT_max * 10 / self.targetExitValue_Nom + 45) - self.targetTrueValue_Cur) / self.targetTrueValue_Nom * 100
				calc_Min = ((OUT_min * 10 / self.targetExitValue_Nom + 45) - self.targetTrueValue_Cur) / self.targetTrueValue_Nom * 100					
			else:
				calc_Max = ((OUT_max * self.targetTrueValue_Nom / self.targetExitValue_Nom) - self.targetTrueValue_Cur) / self.targetTrueValue_Nom * 100
				calc_Min = ((OUT_min * self.targetTrueValue_Nom / self.targetExitValue_Nom) - self.targetTrueValue_Cur) / self.targetTrueValue_Nom * 100

			#if ((abs(calc_Max) > self.targetAccuracy) or (abs(calc_Min) > self.targetAccuracy)):
				#self.outBox.append('БРАК')

			if ((abs(calc_Max)) >= (abs(calc_Min))):
				return OUT_max
			else:
				return OUT_min

	def prot(self, case, sheet_col_init, sheet_row_init):
		# Подготовка протокола
		if case=='init':
			global wb
			global sheet
			global sheet_col
			global sheet_row

			sheet_col = sheet_col_init
			sheet_row = sheet_row_init
			wb = openpyxl.load_workbook(filename="Libary/"+self.targetType_Main.currentText().replace("/","#")+'_'+self.targetExitParam.currentText()+".xlsx")
			sheet = wb.worksheets[0]
			
			# Ввод данных протокола
			sheet['L1'] = self.protNumber_1.text() + self.protNumber_2.text().replace("#","/")
			sheet['N3'] = self.targetType_Sub.text()
			sheet['R3'] = self.targetNumber.text()
			sheet['B20'] = self.targetExitNumber.checkedButton().text()
			sheet['E20'] = self.targetExitParam.currentText()
			sheet['H20'] = self.targetExitCode.currentText()
			sheet['I5'] = self.targetNomVar.text().replace(".",",")
			sheet['T63'] = time.strftime("%d.%m.%Y", time.gmtime())
			
		# Ввод данных поверки
		elif case=='data':
			OUT_series = []
			for i in range(int(float(self.measureTime.text()) / config.ser.timeout)):
				config.ser.write(b'MEAS? \n')
				OUT = config.ser.read(17).strip().decode("utf-8") # Создает задержку равную параметру config.ser.timeout

				# Преобразование выходных данных
				OUT_point = OUT.find('e')
				EXP = OUT[OUT_point+1:]
				EXP = int(EXP)+3
				OUT = float(OUT[:6])*10**EXP
				OUT_series.append(OUT)
				print(OUT)

			OUT = self.accuracyCalc(OUT_series)

			# Занос в протокол
			OUT = str(OUT)
			OUT = OUT[:6].replace(".",",")
			self.outBox.append(OUT)
			sheet[sheet_col+str(sheet_row)] = OUT
			sheet_row = int(sheet_row)
			sheet_row += 1
		
		# Вывод протокола
		elif case == 'end':
			try:
				int(self.targetExitNumber.checkedButton().text())
			except:
				wb.save(filename="Protocols/"+self.protNumber_1.text()+self.protNumber_2.text()+".xlsx")
			else:
				wb.save(filename="Protocols/"+self.protNumber_1.text()+"-"+self.targetExitNumber.checkedButton().text()+self.protNumber_2.text()+".xlsx")

			self.outBox.append("Повірку завершено")
	
	def warmup(self, mode_warmup, perc_warmup, nom_var):
		self.outBox.append("Прогрев")
		# Подготовка данных
		mode_warmup = str(mode_warmup).encode('ascii')
		var_warmup = str(nom_var*perc_warmup).encode('ascii')
			
		# Прогрев
		config.ser.write(mode_warmup+b" "+var_warmup+b" \n")
		config.ser.write(b"OUTP ON \n")
		self.flagSleep(self.time_warmup)
		self.outBox.append("Повірка:")

	#-----------
	# Метод работы
	#-----------
	def AC_current(self, perc_warmup, I_nom, I_steps):
	
		# Прогрев
		self.warmup("CAC:CURR", perc_warmup, I_nom)
	
		# Поверка
		i = 0
		while( i < len(I_steps)):
			self.targetTrueValue_Cur = I_steps[i]

			config.ser.write(b'CAC:CURR '+(str(I_steps[i]).encode('ascii'))+b' \n')
			
			self.flagSleep(self.time_step)
			self.prot('data',0,0)
			self.flagSleep(2)
			
			i += 1
			
		self.prot('end',0,0)
		self.stop()

	def AC_voltage(self, case, perc_warmup, U_nom, U_steps, F_nom, F_steps):
		
		# Установка дефолтних значений и настроек
		config.ser.write(b"CONF CURR \n")
		config.ser.write(b"VAC:VOLT "+(str(U_nom).encode('ascii'))+b" \n")
		config.ser.write(b"VAC:FREQ "+(str(F_nom).encode('ascii'))+b" \n")	
	
		# Прогрев
		if (self.targetExitParam.currentText() == 'U'):
			self.warmup("VAC:VOLT", perc_warmup, U_nom)
		
		elif (self.targetExitParam.currentText() == 'F'):
			self.warmup("VAC:FREQ", perc_warmup, F_nom)
			
		# Поверка
		i = 0
		while( i < len(U_steps)):
			# Считаем действительное значение
			if (self.targetExitParam.currentText() == 'U'):
				self.targetTrueValue_Cur = U_steps[i]
			elif (self.targetExitParam.currentText() == 'F'):
				self.targetTrueValue_Cur = F_steps[i]

			config.ser.write(b'VAC:VOLT '+(str(U_steps[i]).encode('ascii'))+b' \n')
			config.ser.write(b'VAC:FREQ '+(str(F_steps[i]).encode('ascii'))+b' \n')
			
			# Перезапуск подачи при переходе от значений меньших за 100 V к большим
			if ((U_steps[i] >= 100) and (U_steps[i-1] < U_steps[i])):
				config.ser.write(b"OUTP ON \n")
			
			self.flagSleep(self.time_step)
			self.prot('data',0,0)
			self.flagSleep(2)
			
			i += 1
			
		self.prot('end',0,0)
		self.stop()
		
	def Power(self, perc_warmup, I_nom, I_steps, U_steps, D_steps):
	
		# Прогрев
		self.warmup("PAC:CURR", perc_warmup, I_nom)

		# Поверка
		i = 0
		while( i < len(I_steps)):
			# Считаем действительное значение
			if (self.targetExitParam.currentText() == 'P'):
				self.targetTrueValue_Cur = U_steps[i] * math.sqrt(3) * I_steps[i] * math.cos(D_steps[i])
			elif (self.targetExitParam.currentText() == 'Q'):
				self.targetTrueValue_Cur = U_steps[i] * math.sqrt(3) * I_steps[i] * math.sin(D_steps[i])
			
			config.ser.write(b'PAC:CURR '+(str(I_steps[i]).encode('ascii'))+b' \n')
			config.ser.write(b'PAC:VOLT '+(str(U_steps[i]).encode('ascii'))+b' \n')
			config.ser.write(b'PAC:PHAS '+(str(D_steps[i]).encode('ascii'))+b' \n')
						
			self.flagSleep(self.time_step)
			self.prot('data',0,0)
			self.flagSleep(2)
			
			i += 1
			
		self.prot('end',0,0)
		self.stop()
		
	#-----------
	# Приборы
	#----------

	def E842_I(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
		perc_warmup = 0.8

		self.targetTrueValue_Nom = I_nom
		
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
		self.AC_current(perc_warmup, I_nom, I_steps)		
		
	def MTE_121_I(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
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
		self.AC_current(perc_warmup, I_nom, I_steps)
		
	def MTE_143_P(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		self.targetTrueValue_Nom = U_nom * math.sqrt(3) * I_nom * math.cos(0)
		
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
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)			

	def MTE_143_Q(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
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
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)
		
	def MTE_143_I(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
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
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)

	def MTE_142_P(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
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
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)

	def MTE_142_Q(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
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
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)				
		
	def MTE_142_I(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
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
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)	
	
	def MTE_111_U(self):
		# Вводные
		U_nom = float(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 0.8

		self.targetTrueValue_Nom = U_nom
	
		#Шаги
		U_steps = [];
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.5)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		U_steps.append (U_nom * 1.1)
		F_steps = [F_nom, F_nom, F_nom, F_nom, F_nom];

		self.prot('init','H','23')		
		self.AC_voltage('volt', perc_warmup, U_nom, U_steps, F_nom, F_steps)	
		
	def MTE_111_F(self):
		# Вводные
		U_nom = float(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 1.06

		self.targetTrueValue_Nom = F_nom
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [45, 47.5, 50, 52.5, 55];

		self.prot('init','E','22')		
		self.AC_voltage('freq', perc_warmup, U_nom, U_steps, F_nom, F_steps)	

	def E858_1_F(self):
		# Вводные
		U_nom = int(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 1
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [45, 45.01, 47.5, 50, 52.5, 54.99, 55];

		self.prot('init','F','22')		
		self.AC_voltage('freq', perc_warmup, U_nom, U_steps, F_nom, F_steps)		
		
	def E858_2_F(self):
		# Вводные
		U_nom = int(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 1
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [48, 49, 50, 51, 52];

		self.prot('init','F','22')		
		self.AC_voltage('freq', perc_warmup, U_nom, U_steps, F_nom, F_steps)			
		
	def E855_1_U(self):
		# Вводные
		U_nom = int(self.targetNomVar.text())
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
		self.AC_voltage('volt', perc_warmup, U_nom, U_steps, F_nom, F_steps)

	def E855_2_U(self):
		# Вводные
		U_nom = int(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 0.92
	
		#Шаги
		U_steps = [75, 85, 95, 105, 115, 125];
		F_steps = [F_nom, F_nom, F_nom, F_nom, F_nom, F_nom];

		self.prot('init','H','23')		
		self.AC_voltage('volt', perc_warmup, U_nom, U_steps, F_nom, F_steps)			
	
	def E848_P(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
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
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)		

	E859_P = E848_P	
		
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = main('gui.ui')
    sys.exit(app.exec_())

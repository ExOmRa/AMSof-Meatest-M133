import openpyxl
import sys
import time
import math
import serial

import config

from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from PySide2.QtUiTools import QUiLoader
from PySide2.QtWidgets import QApplication, QPushButton, QLineEdit, QComboBox, QTextBrowser, QButtonGroup, QLabel, QSpinBox
from PySide2.QtCore import QFile, QObject



class main(QObject):
	#-----------
	# Основные переменные
	#----------- 			
	executor = ThreadPoolExecutor(max_workers = 2)
	statusCode = 'stop'

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
		self.outBoxOne = self.window.findChild(QTextBrowser, 'outBoxOne')
		self.outBoxTwo = self.window.findChild(QTextBrowser, 'outBoxTwo')
		#self.outBoxThree = self.window.findChild(QTextEdit, 'outBoxThree')

		# Настройки связи
		self.comNumber = self.window.findChild(QLineEdit, 'comNumber')
		self.comFreq = self.window.findChild(QLineEdit, 'comFreq')
		
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
		self.protNumber_1 = self.window.findChild(QSpinBox, 'protNumber_1')
		self.protNumber_2 = self.window.findChild(QLineEdit, 'protNumber_2')
		self.time_warmup_Link = self.window.findChild(QLineEdit, 'warmupTime')
		self.time_step_Link = self.window.findChild(QLineEdit, 'oneStepTime')
		self.measureTime = self.window.findChild(QLineEdit, 'measureTime')

		# Управление
		self.window.findChild(QPushButton, 'startButton').clicked.connect(self.start_button)
		self.window.findChild(QPushButton, 'stopButton').clicked.connect(self.stop_button)
		self.window.findChild(QPushButton, 'testButton').clicked.connect(self.test_button)

		# Обратная связь
		self.diagnosLabel = self.window.findChild(QLabel, 'diagnosLabel')
		
		# События формы
		self.window.show()	

	#-----------
	# Обработка событий формы
	#-----------
	def start_button(self):
		if self.statusCode == 'stop':
			self.statusCode = 'working'
			self.connect()
			self.outBoxOne.clear()
			self.outBoxTwo.clear()

			try: 
				device_start_execute = getattr(self, self.targetType_Main.currentText().replace("/","_")+'_'+self.targetExitParam.currentText())
			except:
				self.statusCode = 'stop'
				self.outBoxOne.append('- Відсутній набір повірочних значень для даного коду виходу')
			else:
				# Общие параметры
				self.diagnos = "Придатний"
				self.stepNumber = 0
				self.targetAccuracy = float(self.targetAccuracy_Link.text().replace(",","."))
				self.time_warmup = int(self.time_warmup_Link.text())
				self.time_step = int(self.time_step_Link.text())			
				self.diagnosLabel.setText('. . .')
				if (self.targetExitCode.currentText() == '0..20'):
					self.targetExitValue_Nom = 20
				else:
					self.targetExitValue_Nom = 5
				
				# Запуск
				self.ser.write(b'SYST:REM \n')
				self.executor.submit(device_start_execute)

	def stop_button(self):
		self.stop()

	def test_button(self):
		self.outBoxOne.clear()
		self.connect()

	#-----------
	# Методы
	#-----------
	def connect(self):
		try: 
			self.ser = serial.Serial(port='COM' + self.comNumber.text(), timeout=float(self.comFreq.text()) * 0.001, write_timeout = 0.1)
			self.ser.write(b'SYST:REM \n')
			self.ser.write(b'MEAS? \n')
			OUT = self.ser.read(17).strip().decode("utf-8") # Создает задержку равную параметру self.ser.timeout
			self.ser.write(b"SYST:LOC \n")
		except:
			self.outBoxOne.append("- Помилка з'єднання - Порт відсутній")
			#self.stop()
			#self.ser.write(b"SYST:LOC \n")
			#self.ser.close()
		else:
			if OUT == '':
				self.outBoxOne.append("- Помилка з'єднання - Порт зайнято іншим пристроем")
				self.stop()
				#self.ser.write(b"SYST:LOC \n")
				#self.ser.close()
			elif (self.statusCode != 'working'):
				self.outBoxOne.append("- OK - З'єднання наявне - OK -")		
				self.stop()
				#self.ser.write(b"SYST:LOC \n")
				#self.ser.close()

	def stop(self):
		self.ser.write(b"OUTP OFF \n")
		self.ser.write(b"SYST:LOC \n")
		# Виводим протокол якщо є хоть 1 отримане значення 
		if ((self.statusCode == 'working') and (self.stepNumber > 0)):
			self.prot('end',0,0)
		self.statusCode = 'stop'	
		self.ser.close()

	def flagSleep(self, timeToSleep):
		for i in range(timeToSleep):
			if self.statusCode == 'stop':
				sys.exit()
			time.sleep(1)
			
	def readData(self):
		self.ser.write(b'MEAS? \n')
		OUT = self.ser.read(17).strip().decode("utf-8") # Создает задержку равную параметру self.ser.timeout
		# Преобразование входных данных
		OUT_point = OUT.find('e')
		EXP = OUT[OUT_point+1:]
		EXP = int(EXP)+3
		OUT = float(OUT[:6])*10**EXP
		return OUT
		
	def toOutput(self, case, val):
		val = str(('{:.4f}'.format(round(val,4)))).replace(".",",")
	
		if case == 'box1':
			self.outBoxOne.append("№" + str(self.stepNumber + 1) + " : " + val)
			
		elif case == 'box2':
			self.outBoxTwo.append(val)
			self.outBoxTwo.ensureCursorVisible()
		
	def accuracyCalc(self, OUT_series):
		# Убираем минус из расчетного если прибор не имеет минусового выхода
		if (('-' not in self.targetExitCode.currentText()) and (self.targetTrueValue_Cur < 0)):
			self.targetTrueValue_Cur *= -1

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

		if ((abs(calc_Max) > self.targetAccuracy) or (abs(calc_Min) > self.targetAccuracy)):
			self.diagnos = 'Брак'
			self.diagnosLabel.setText(self.diagnos)
			self.outBoxOne.append('БРАК')

			#self.outBoxOne.append(str(self.targetTrueValue_Nom))
			#self.outBoxOne.append(str(self.targetTrueValue_Cur))

		if ((abs(calc_Max)) >= (abs(calc_Min))):
			return OUT_max
		else:
			return OUT_min

	def prot(self, case, sheetProtColStart, sheetProtRowStart):
		# Подготовка протокола
		if case=='init':
			global wb
			wb = openpyxl.load_workbook(filename="Libary/"+self.targetType_Main.currentText().replace("/","#")+'_'+self.targetExitParam.currentText()+".xlsx")
			
			global sheetProt
			global sheetProtCol
			global sheetProtRow

			sheetProt = wb.worksheets[0]
			sheetProtCol = sheetProtColStart
			sheetProtRow = sheetProtRowStart			
			
			global sheetSig
			global sheetSigCol
			global sheetSigRow			
			
			# Parameters for signal sheet
			sheetSig = wb.worksheets[1]
			sheetSigCol = 'E'
			sheetSigRow = '1'
			
			# Parameters for protocol sheet
			sheetProt['L1'] = str(self.protNumber_1.value()) + self.protNumber_2.text().replace("#","/")
			sheetProt['N3'] = self.targetType_Sub.text()
			sheetProt['R3'] = self.targetNumber.text()
			sheetProt['B20'] = self.targetExitNumber.checkedButton().text()
			sheetProt['E20'] = self.targetExitParam.currentText()
			sheetProt['H20'] = self.targetExitCode.currentText()
			sheetProt['I5'] = self.targetNomVar.text().replace(".",",")
			sheetProt['T63'] = time.strftime("%d.%m.%Y", time.gmtime())
			
		# Processing data
		elif case=='data':
			OUT_series = []
			for i in range(int(float(self.measureTime.text()) / self.ser.timeout)):
				OUT = self.readData()

				# Output data for signal sheet
				sheetSig[sheetSigCol+str(sheetSigRow)] = OUT
				sheetSigRow = int(sheetSigRow)
				sheetSigRow += 1				

				self.toOutput('box2', OUT)
				OUT_series.append(OUT)	

			OUT = self.accuracyCalc(OUT_series)
			self.toOutput('box1', OUT)
			
			# Output data for protocol sheet
			sheetProt[sheetProtCol+str(sheetProtRow)] = OUT
			sheetProtRow = int(sheetProtRow)
			sheetProtRow += 1
			
		# Protocol output
		elif case == 'end':
			sheetProt['D63'] = self.diagnos
			self.diagnosLabel.setText(self.diagnos)
			
			try:
				int(self.targetExitNumber.checkedButton().text())
			except:
				protName = "Protocols/"+str(self.protNumber_1.value())+self.protNumber_2.text()+"-"+self.targetNumber.text()
			else:
				protName = "Protocols/"+str(self.protNumber_1.value())+"-"+self.targetExitNumber.checkedButton().text()+self.protNumber_2.text()+"-"+self.targetNumber.text()

			# Checking for double of protocol file name
			if Path(protName+".xlsx").is_file():
				wb.save(filename=protName + "_time(" + time.strftime("%H.%M", time.gmtime()) + ").xlsx")
				self.outBoxOne.append("- Протокол з такою назвою вже існує, до назви поточного додано мітку часу")
			else:
				wb.save(filename=protName + ".xlsx")
				
			self.outBoxOne.append("- Повірку завершено")
	
	def warmup(self, mode_warmup, perc_warmup, nom_var):
		self.outBoxOne.append("- Виконується прогрів приладу")
		# Подготовка данных
		mode_warmup = str(mode_warmup).encode('ascii')
		var_warmup = str(nom_var*perc_warmup).encode('ascii')
			
		# Прогрев
		self.ser.write(mode_warmup+b" "+var_warmup+b" \n")
		self.ser.write(b"OUTP ON \n")
		self.flagSleep(self.time_warmup)
		self.outBoxOne.append("- Початок повірки [№ точки : значення з макс. похибкою]:")

	#-----------
	# Метод работы
	#-----------
	def AC_current(self, perc_warmup, I_nom, I_steps):
	
		# Прогрев
		self.warmup("CAC:CURR", perc_warmup, I_nom)
	
		# Поверка
		while(self.stepNumber < len(I_steps)):
			self.targetTrueValue_Cur = I_steps[self.stepNumber]

			self.ser.write(b'CAC:CURR '+(str(I_steps[self.stepNumber]).encode('ascii'))+b' \n')
			
			self.flagSleep(self.time_step)
			self.prot('data',0,0)
			self.flagSleep(2)
			
			self.stepNumber += 1
			
		self.stop()
			
	def AC_voltage(self, perc_warmup, U_nom, U_steps, F_nom, F_steps):
		
		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"VAC:VOLT "+(str(U_nom).encode('ascii'))+b" \n")
		self.ser.write(b"VAC:FREQ "+(str(F_nom).encode('ascii'))+b" \n")	
	
		# Прогрев
		if (self.targetExitParam.currentText() == 'U'):
			self.warmup("VAC:VOLT", perc_warmup, U_nom)
		
		elif (self.targetExitParam.currentText() == 'F'):
			self.warmup("VAC:FREQ", perc_warmup, F_nom)
			
		# Поверка
		while(self.stepNumber < len(U_steps)):
			# Считаем действительное значение
			if (self.targetExitParam.currentText() == 'U'):
				self.targetTrueValue_Cur = U_steps[self.stepNumber]
			elif (self.targetExitParam.currentText() == 'F'):
				self.targetTrueValue_Cur = F_steps[self.stepNumber]

			self.ser.write(b'VAC:VOLT '+(str(U_steps[self.stepNumber]).encode('ascii'))+b' \n')
			self.ser.write(b'VAC:FREQ '+(str(F_steps[self.stepNumber]).encode('ascii'))+b' \n')
			
			# Перезапуск подачи при переходе от значений меньших за 100 V к большим
			if ((U_steps[self.stepNumber] >= 100) and (U_steps[self.stepNumber-1] < U_steps[self.stepNumber])):
				self.ser.write(b"OUTP ON \n")
			
			self.flagSleep(self.time_step)
			self.prot('data',0,0)
			self.flagSleep(2)
			
			self.stepNumber += 1
			
		self.stop()
			
	def Power(self, perc_warmup, I_nom, I_steps, U_steps, D_steps):
	
		# Прогрев
		self.warmup("PAC:CURR", perc_warmup, I_nom)

		# Поверка
		while( self.stepNumber < len(I_steps)):
			# Считаем действительное значение
			if (self.targetExitParam.currentText() == 'P'):
				self.targetTrueValue_Cur = U_steps[self.stepNumber] * 3 * I_steps[self.stepNumber] * math.cos(math.radians(D_steps[self.stepNumber]))
			elif (self.targetExitParam.currentText() == 'Q'):
				self.targetTrueValue_Cur = U_steps[self.stepNumber] * 3 * I_steps[self.stepNumber] * math.sin(math.radians(D_steps[self.stepNumber]))
			elif (self.targetExitParam.currentText() == 'I'):
				self.targetTrueValue_Cur = I_steps[self.stepNumber]

			self.ser.write(b'PAC:CURR '+(str(I_steps[self.stepNumber]).encode('ascii'))+b' \n')
			self.ser.write(b'PAC:VOLT '+(str(U_steps[self.stepNumber]).encode('ascii'))+b' \n')
			self.ser.write(b'PAC:PHAS '+(str(D_steps[self.stepNumber]).encode('ascii'))+b' \n')
						
			self.flagSleep(self.time_step)
			self.prot('data',0,0)
			self.flagSleep(2)
			
			self.stepNumber += 1
			
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
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"CAC:FREQ 50 \n")
		
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

		self.targetTrueValue_Nom = I_nom

		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"CAC:FREQ 50 \n")
		
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

		self.targetTrueValue_Nom = U_nom * 3 * I_nom * math.cos(math.radians(0))
		
		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"PAC:UNIT W \n")
		self.ser.write(b"PAC:VOLT 57.735 \n")
		self.ser.write(b"PAC:FREQ 50 \n")
		self.ser.write(b"PAC:PHAS 0 \n")
		self.ser.write(b"OUTP:CONF 123 \n")
		
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
		D_steps.append (60)
		D_steps.append (150)
		D_steps.append (240)
		D_steps.append (330)			
		
		self.prot('init','P','23')
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)			

	def MTE_143_Q(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		self.targetTrueValue_Nom = U_nom * 3 * I_nom * math.sin(math.radians(90))

		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"PAC:UNIT VAR \n")
		self.ser.write(b"PAC:VOLT 57.735 \n")
		self.ser.write(b"PAC:FREQ 50 \n")
		self.ser.write(b"PAC:PHAS 90 \n")
		self.ser.write(b"OUTP:CONF 123 \n")
		
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
		D_steps.append (180)
		D_steps.append (90)
		D_steps.append (270)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (60)
		D_steps.append (150)
		D_steps.append (120)
		D_steps.append (330)
		
		self.prot('init','P','23')		
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)
		
	def MTE_143_I(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		self.targetTrueValue_Nom = I_nom

		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"PAC:UNIT W \n")
		self.ser.write(b"PAC:VOLT 57.735 \n")
		self.ser.write(b"PAC:FREQ 50 \n")
		self.ser.write(b"PAC:PHAS 0 \n")
		self.ser.write(b"OUTP:CONF 123 \n")
		
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

		self.targetTrueValue_Nom = U_nom * 3 * I_nom * math.cos(math.radians(0))

		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"PAC:UNIT W \n")
		self.ser.write(b"PAC:VOLT 57.735 \n")
		self.ser.write(b"PAC:FREQ 50 \n")
		self.ser.write(b"PAC:PHAS 0 \n")
		self.ser.write(b"OUTP:CONF 123 \n")
		
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
		D_steps.append (60)
		D_steps.append (150)
		D_steps.append (240)
		D_steps.append (330)	
		
		self.prot('init','P','23')
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)

	def MTE_142_Q(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		self.targetTrueValue_Nom = U_nom * 3 * I_nom * math.sin(math.radians(90))

		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"PAC:UNIT VAR \n")
		self.ser.write(b"PAC:VOLT 57.735 \n")
		self.ser.write(b"PAC:FREQ 50 \n")
		self.ser.write(b"PAC:PHAS 90 \n")
		self.ser.write(b"OUTP:CONF 123 \n")
		
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
		D_steps.append (180)
		D_steps.append (90)		
		D_steps.append (270)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (90)
		D_steps.append (60)
		D_steps.append (150)
		D_steps.append (120)
		D_steps.append (330)
		
		self.prot('init','P','23')
		self.Power(perc_warmup, I_nom, I_steps, U_steps, D_steps)				
		
	def MTE_142_I(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		self.targetTrueValue_Nom = I_nom

		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"PAC:UNIT W \n")
		self.ser.write(b"PAC:VOLT 57.735 \n")
		self.ser.write(b"PAC:FREQ 50 \n")
		self.ser.write(b"PAC:PHAS 0 \n")
		self.ser.write(b"OUTP:CONF 123 \n")
		
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
		self.AC_voltage(perc_warmup, U_nom, U_steps, F_nom, F_steps)	
		
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
		self.AC_voltage(perc_warmup, U_nom, U_steps, F_nom, F_steps)	

	def E858_1_F(self):
		# Вводные
		U_nom = float(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 1

		self.targetTrueValue_Nom = F_nom
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [45, 45.01, 47.5, 50, 52.5, 54.99, 55];

		self.prot('init','F','22')		
		self.AC_voltage(perc_warmup, U_nom, U_steps, F_nom, F_steps)		
		
	def E858_2_F(self):
		# Вводные
		U_nom = float(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 1

		self.targetTrueValue_Nom = F_nom
		
		# Шаги
		U_steps = [U_nom, U_nom, U_nom, U_nom, U_nom];
		F_steps = [48, 49, 50, 51, 52];

		self.prot('init','F','22')		
		self.AC_voltage(perc_warmup, U_nom, U_steps, F_nom, F_steps)			
		
	def E855_1_U(self):
		# Вводные
		U_nom = float(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 0.8

		self.targetTrueValue_Nom = U_nom
	
		#Шаги
		U_steps = [];
		U_steps.append (U_nom * 0.2)
		U_steps.append (U_nom * 0.4)
		U_steps.append (U_nom * 0.6)
		U_steps.append (U_nom * 0.8)
		U_steps.append (U_nom * 1)
		F_steps = [F_nom, F_nom, F_nom, F_nom, F_nom];

		self.prot('init','H','23')		
		self.AC_voltage(perc_warmup, U_nom, U_steps, F_nom, F_steps)

	def E855_2_U(self):
		# Вводные
		U_nom = float(self.targetNomVar.text())
		F_nom = 50
		perc_warmup = 0.92

		self.targetTrueValue_Nom = U_nom
	
		#Шаги
		U_steps = [75, 85, 95, 105, 115, 125];
		F_steps = [F_nom, F_nom, F_nom, F_nom, F_nom, F_nom];

		self.prot('init','H','23')		
		self.AC_voltage(perc_warmup, U_nom, U_steps, F_nom, F_steps)			
	
	def E848_P(self):
		# Вводные
		I_nom = float(self.targetNomVar.text())
		U_nom = 57.735
		perc_warmup = 0.8

		self.targetTrueValue_Nom = U_nom * 3 * I_nom * math.cos(math.radians(0))

		# Установка дефолтних значений и настроек
		self.ser.write(b"CONF CURR \n")
		self.ser.write(b"PAC:UNIT W \n")
		self.ser.write(b"PAC:VOLT 57.735 \n")
		self.ser.write(b"PAC:FREQ 50 \n")
		self.ser.write(b"PAC:PHAS 0 \n")
		self.ser.write(b"OUTP:CONF 123 \n")
		
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
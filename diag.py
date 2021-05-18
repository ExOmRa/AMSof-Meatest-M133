# 1Установку дефолтних значений и настроек в методы работы
# Для MTE 143 нет 4 проводной

import openpyxl
import sys
import time
import math
import serial

ser = serial.Serial(port='COM1', timeout=250)
ser.write(b'SYST:REM \n')
ser.write(b"OUTP ON \n")
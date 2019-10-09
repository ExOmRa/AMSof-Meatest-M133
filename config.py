import serial
import sys

sys.path.insert(0, 'Libary/')
ser = serial.Serial(port='COM1',timeout=2)
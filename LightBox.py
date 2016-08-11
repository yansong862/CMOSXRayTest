import serial
import time
import threading

class LightBox:
    
    def __init__(self, portNum):
        self.ackRecieved = False
        self.stopThread = False
        self.ser = serial.Serial()
        self.ser.port = portNum
        self.ser.baudrate = 9600
        self.ser.bytesize = serial.EIGHTBITS
        self.ser.parity = serial.PARITY_NONE
        self.ser.stopbits = serial.STOPBITS_ONE 
        self.ser.timeout = 1 
        self.ser.xonxoff = False 
        self.ser.rtscts = False 
        self.ser.dsrdtr = False 
        self.ser.writeTimeout = 2  
        self.readBuff = ''
        self.tempBuff = ''
        
    def SetIntensity(self, intensity):
        if(intensity<0):
            intensity = 0
        if(intensity>4095):
            intensity = 4095
        self.ser.open()
        command = "L" + str(intensity).zfill(4) + "\r"
        print("Command asdf: " + command)
        if self.ser.isOpen():
            self.ser.write(command)
            time.sleep(0.1)
        self.ser.close()

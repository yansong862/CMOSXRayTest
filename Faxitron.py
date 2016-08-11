import serial
import time
import threading

class Cabinet:
    'Faxitron cabinet communications class'

    def SerialRead(self):
        while (self.stopThread == False):    
            while self.ser.inWaiting() > 0:
                self.tempBuff += self.ser.read(1)
            if self.tempBuff != '':
                self.ser.flushInput()
                #print ">>" + self.tempBuff
                if(self.tempBuff == '\x1BTMSS0000800008\x0D'):
                    print("Cabinet Door Open!")
                elif(self.tempBuff == '\x1BTMSS0000000000\x0D'):
                    print("Cabinet Door Closed!")
                elif(self.tempBuff == '\x1B\x06\x0D'):
                    #print("Ack Received!")
                    self.ackRecieved = True
                else:
                    self.readBuff = self.tempBuff
                self.tempBuff = ''
        self.readBuff = ''
        
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
        try: 
            self.ser.open()
        except  e:
            print ("error open serial port: " + str(e))
            exit()
            
        t = threading.Thread(target=self.SerialRead)
        t.start()
    
    def __del__(self):
        self.stopThread = True
        if self.ser.isOpen():
            self.ser.close()
    
    def close(self):
        self.stopThread = True
        self.ser.close()
            
    def SendCommand(self,command):
        counter = 0
        ackRecieved = False
        fullCommand = '\x1B' + command + '\x0D'
        if self.ser.isOpen():
            self.ser.write(fullCommand)
            time.sleep(0.1) 
            #self.readBuff = ''
        else:
            print ("cannot open serial port ")
        
    def Configure(self,kV,mS):
        counter = 0
        self.ackRecieved = False
        _kV = "%.2f" % kV
        mA = 1
        _mA = "%.1f" % mA
        command = 'CBM' + str(_kV).zfill(6)+str(_mA).zfill(6)+str(mS).zfill(6)
        self.SendCommand(command)
        while(counter < 5 and self.ackRecieved == False):
            time.sleep(0.1)
            counter += 1
        if self.ackRecieved == True:
            return True
        else: 
            return False
        
    def GetConfig(self):
        command = 'RB'
        self.SendCommand(command)
        time.sleep(0.1)
        if(self.readBuff != ''):
            kv = float(self.readBuff[7:13])
            #print("KV: %.2f KV"  % kv)
            mA = float(self.readBuff[13:19])
            #print("mA: %.1f mA"  % mA)
            mS = int(self.readBuff[19:-1])
            #print("Duration: %d mS"  % mS)
            return (kv,mA,mS)
            
    def SetKV(self, kV):
        mS = 0
        command = 'RB'
        self.SendCommand(command)
        time.sleep(0.1)
        if(self.readBuff != ''):
            mS = int(self.readBuff[19:-1])
        return self.Configure(kV,mS)
        
    def GetKV(self):
        command = 'RK'
        self.SendCommand(command)
        time.sleep(0.1)
        if(self.readBuff != ''):
            kv = float(self.readBuff[3:-1])
            #print("KV is: %.2f KV" % kv)
            return kv            
    
    def SetTime(self, mS):
        kV = 0
        command = 'RB'
        self.SendCommand(command)
        time.sleep(0.1)
        if(self.readBuff != ''):
            kV = float(self.readBuff[7:13])
        return self.Configure(kV,mS)
        
    def GetTime(self):
        command = 'RT'
        self.SendCommand(command)
        time.sleep(0.1)
        if(self.readBuff != ''):
            mS = int(self.readBuff[3:-1])
            #print("Duration is: %d ms" % mS)
            return mS
        
    def FireXRay(self):
        counter = 0
        self.ackRecieved = False
        self.SendCommand('CMSS0006400064')
        time.sleep(0.1)
        self.SendCommand('CMSS0001600000')
        time.sleep(0.1)
        self.SendCommand('CX1')
        time.sleep(0.1)
        while(counter < 5 and self.ackRecieved == False):
            time.sleep(0.1)
            counter += 1
        if self.ackRecieved == True:
            return True
        else: 
            return False
    
    def WaitXRayOff(self):
        counter = 0
        print("\nX-Ray off"),
        while self.readBuff == '\x1BTDTR1\x0D' and counter <5000:
            print("-"),
            time.sleep(0.01)
            counter +=1
        print("")
            
    def WaitXRayOn(self):
        counter = 0
        print("\nX-Ray on"),
        while self.readBuff != '\x1BTDTR1\x0D' and counter <5000:
            print("+"),
            time.sleep(0.001)
            counter += 1
        print(counter)

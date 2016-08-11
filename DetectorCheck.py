import DexelaPy
import Utils
import DexList
#import Faxitron
import time
import copy
import winsound
import easygui
#import LightBox
import os
import sys

        
def DetCheck(det,portNum):
    
    print("Connecting to detector...")
    det.OpenBoard()
    print ("done OpenBoard")
    det.SetExposureMode(DexelaPy.ExposureModes.Sequence_Exposure)
    print ("done SetExps...")
    model = det.GetModelNumber()
    serial = det.GetSerialNumber()
    filename = str(model)+"-"+str(serial)+"-"
    outpath = "Output\\"
    #outpath = "C:\\Users\\plagon\\Documents\\Projects\\test\\Array_1207_454\\X-Ray\\"
    
    counter = 1
    print("Calibrating ADC Offsets!")
    tempFileName = outpath+"AutoCal\\" + filename + "_autocal.txt"
    print "file name :", tempFileName
    darkOffset = int(300)
    StopDarkCalibration = 15000
    StartDarkCalibration = 14000
    ADCBoardType = easygui.buttonbox('Select ADC Board Type', 'ADC Board Type', ('ADC','ADC-RP'))
    if  ADCBoardType == 'ADC':
        StopDarkCalibration = 15000
        StartDarkCalibration = 14000
    else:
        StopDarkCalibration = 5100
        StartDarkCalibration = 4800

    print(ADCBoardType, StopDarkCalibration, StartDarkCalibration, darkOffset)
    Utils.GetDarkOffsets(det,tempFileName,darkOffset,StopDarkCalibration,StartDarkCalibration)
    #Utils.GetDarkOffsets(det,tempFileName)
    print("\n")
        
    print("Image successfully saved!")
    det.CloseBoard()
        
    return 

try:
    
    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    for i in range(0,count):
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        DetCheck(det, 3)

except DexelaPy.DexelaExceptionPy as ex:
    print("Exception Occurred!")
    print("Description: %s" % ex)
    DexException = ex.DexelaException
    print("Function: %s" % DexException.GetFunctionName())
    print("Line number: %d" % DexException.GetLineNumber())
    print("File name: %s" % DexException.GetFileName())
    print("Transport Message: %s" % DexException.GetTransportMessage())
    print("Transport Number: %d" % DexException.GetTransportError())
except Exception:
    print("Exception OCCURRED!")




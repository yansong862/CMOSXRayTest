from ctypes import *
import time
import copy
import winsound
import easygui
import os
import sys
import glob
import filecmp
import ntpath
import datetime

from numpy.lib.scimath import logn
from math import e



import DexelaPy 




def runMeasTemperature():
    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    measedT =0.0
    
    if count>0:
        i =0
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        det.OpenBoard()

        ### detector info
        model = det.GetModelNumber()
        serial = det.GetSerialNumber()
        firmVer =det.GetFirmwareVersion()

        Rt =8.25
        B =3950.0
        Tr=298.15
        T0 =273.15
        det.WriteRegister(116,3) #enable temperature read
        vadc =det.ReadRegister(112)
        print("Vadc: %d " % vadc)
        r =Rt/(4095/(vadc*1.0)-1)
        invertT =1/Tr+(1/B)*logn(e, r/10)
        measedT =1.0/invertT-T0
        print (str("%f" %measedT))
        
        det.CloseBoard()

        print("Detector" + str(model)+"-"+str(serial)+" temperature " +str("%f" %measedT) + " (register value " +str("%d " % vadc)+")")

    return measedT


### runMeasTemperature()

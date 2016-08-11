#import DexelaPy
import Utils
import DexList
import Faxitron
import time
import copy
import winsound
import easygui
#import LightBox
import os
import sys

        
        
def dosimeas(cab,element):
    
    print("Dosi meas. sequence with:")
    print("KV: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    
    if(element.kVp != 0):
        cab.Configure(element.kVp,element.xOnTimeSecs*1000)
        if cab.FireXRay() != True:
            return -1
        cab.WaitXRayOn()
        time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
        cab.WaitXRayOff()
        
        
          
        
def RunScript(list,portNum):
    global cab    
   
    counter = 1
    
    for element in list:
        if element.command == 'generatorinit':
            print("Initializing Generator!")
            cab = Faxitron.Cabinet(portNum)
        elif element.command == 'message':
            easygui.msgbox(element.comment,title="Alert!")
        elif element.command == 'dosimeas':
            dosimeas(cab,element)
            print("\n")
        else:
            print(element.command," is not support in this test")
        
    if cab != None:
        cab.close()
    time.sleep(0.5)
    print "All Done."
    
    return 

try:
    global cab
    cab = None

    cfgFile ='none'
    cfgFile ='C:\CMOS\Configs\cal_seq_FaxitronXRay_dosimeter.xml'
        
    print "loading test sequence from ", cfgFile
    list = DexList.DexList(cfgFile).xmlList
    
    RunScript(list,3)

except Exception:
    print("Exception OCCURRED!")
    #global cab
    if cab != None:
        cab.close()




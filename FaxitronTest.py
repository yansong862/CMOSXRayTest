import DexelaPy
import Utils
import DexList
import Faxitron
import time
import copy
import winsound
import easygui
import LightBox
import os

global defectMaps
defectMaps = ["","",""]
global darkIms
darkIms = ["","","","","",""]
global gainIms
gainIms = ["","","","","",""]

def LoatDefectMaps(path):
    global defectMaps
    global darkIms
    global gainIms
    defectPath = path+"Defect Map\\"
    defectMapFiles = [f for f in os.listdir(defectPath)]
    for f in defectMapFiles:
        if "1x1" in f:
            defectMaps[0] = defectPath+f
        elif "2x2" in f:
            defectMaps[1] = defectPath+f
        elif "4x4" in f:
            defectMaps[2] = defectPath+f
    darkPath = path+"Darks\\"
    darkFiles = [f for f in os.listdir(darkPath)]
    for f in darkFiles:
        if "High" in f:
            if "Binx11" in f:
                darkIms[0] = darkPath+f
            elif "Binx22" in f:
                darkIms[1] = darkPath+f
            elif "Binx44" in f:
                darkIms[2] = darkPath+f
        elif "Low" in f:
            if "Binx11" in f:
                darkIms[3] = darkPath+f
            elif "Binx22" in f:
                darkIms[4] = darkPath+f
            elif "Binx44" in f:
                darkIms[5] = darkPath+f
    gainPath = path+"Floods\\"
    gainFiles = [f for f in os.listdir(gainPath)]
    for f in gainFiles:
        if "High" in f:
            if "Binx11" in f:
                gainIms[0] = gainPath+f
            elif "Binx22" in f:
                gainIms[1] = gainPath+f
            elif "Binx44" in f:
                gainIms[2] = gainPath+f
        elif "Low" in f:
            if "Binx11" in f:
                gainIms[3] = gainPath+f
            elif "Binx22" in f:
                gainIms[4] = gainPath+f
            elif "Binx44" in f:
                gainIms[5] = gainPath+f

    
def TestSequence(outpath,filename,det,cab,element,expNum):
    global defectMaps
    global darkIms
    global gainIms
    
    _t_expms = "%.1f" % element.t_expms
    print("Capturing sequence with:")
    print("KV: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    print("Well-Mode: " + str(element.fullWell))
    print("Binning-Mode: " + str(element.Binning))
    print("Exposure-Time: " + str(element.t_expms))
    print("Number of Exposures: " + str(element.numberOfExposures))
    darkCorrect = element.darkCorrect
    if darkCorrect:
        print("Dark correction is enabled")
    else:
        print("Dark correction is disabled")
    gainCorrect = element.gainCorrect
    if gainCorrect:
        print("Gain correction is enabled")
    else:
        print("Gain correction is disabled")
    defectCorrect = element.defectCorrect
    if defectCorrect:
        print("Defect correction is enabled")
    else:
        print("Defect correction is disabled")
    
    img = DexelaPy.DexImagePy()
    det.SetFullWellMode(element.fullWell)    
    det.SetBinningMode(element.Binning)
    det.SetExposureTime(element.t_expms)
    det.SetTriggerSource(element.trigger)
    det.SetNumOfExposures(element.numberOfExposures)
    filename = filename + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning)

    if gainCorrect:
        filename = filename + "-GainCorrected-"
    
    if defectCorrect:
        filename = filename + "-DefectCorrected-"
    
    det.GoLiveSeq(0,element.numberOfExposures-1,element.numberOfExposures)
    count = det.GetFieldCount()
    endCount = count+element.numberOfExposures+1
    if(element.kVp != 0):
        cab.Configure(element.kVp,element.xOnTimeSecs*1000)
        if cab.FireXRay() != True:
            return -1
        cab.WaitXRayOn()
    
    det.SoftwareTrigger()
    print("Acquiring Image!")
    
    while (count<endCount) :
        count = det.GetFieldCount()
        print("."),
        time.sleep(0.2)
    
    cab.WaitXRayOff()
    print("Image Captured!")
    for i in range (0,element.numberOfExposures):
        det.ReadBuffer(i,img,i)
        
    img.UnscrambleImage()
    
    wellOffset = 0    
    if element.fullWell == DexelaPy.FullWellModes.Low:
        wellOffset =3
    if element.Binning == DexelaPy.bins.x11:
        binOffset = 0
    elif element.Binning == DexelaPy.bins.x22:
        binOffset = 1
    elif element.Binning == DexelaPy.bins.x44:
        binOffset = 2   
    offset = wellOffset+binOffset
    
    if darkCorrect:          
        if os.path.isfile(darkIms[offset]):
            img.LoadDarkImage(darkIms[offset])
            img.SubtractDark()           
            if gainCorrect:
                if os.path.isfile(gainIms[offset]):
                    img.LoadFloodImage(gainIms[offset])
                    img.FloodCorrection()
                else:
                    print("Flood image not found! Cannot perform flood correction!")
        else:
            print("Dark image not found! Cannot perform dark correction!")
    if defectCorrect:
        if os.path.isfile(defectMaps[binOffset]):
            img.LoadDefectMap(defectMaps[binOffset])
            img.DefectCorrection()
        else:
            print("Defect Map not found! Cannot perform defect correction!")        
    
    _kVp = "%.0f" % element.kVp    
    _t_expms = "%.2f" % element.t_expms
    
    if(element.comment != None):
        filename += "_"+str(_kVp)+"kVp_" + str(_t_expms) + "ms_300uA_" + element.comment + "_exp" + str(expNum).zfill(4)
    else:
        filename += "_"+str(_kVp)+"kVp_" + str(_t_expms) + "ms_300uA_exp" + str(expNum).zfill(4)
    
    filename = outpath+filename + ".tif"

    if element.command != 'dummy':
        img.WriteImage(filename)
        
        
        
        
def DarkFloodCapture(outpath,filename,det,cab,element,expNum):
    global darkIms
    global gainIms
    
    _t_expms = "%.1f" % element.t_expms
    print("Capturing sequence with:")
    print("KV: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    print("Well-Mode: " + str(element.fullWell))
    print("Binning-Mode: " + str(element.Binning))
    print("Exposure-Time: " + str(element.t_expms))
    print("Number of Exposures: " + str(element.numberOfExposures))
    
    img = DexelaPy.DexImagePy()
    det.SetFullWellMode(element.fullWell)    
    det.SetBinningMode(element.Binning)
    det.SetExposureTime(element.t_expms)
    det.SetTriggerSource(element.trigger)
    det.SetNumOfExposures(element.numberOfExposures)
    filename = filename + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning)
    det.GoLiveSeq(0,element.numberOfExposures-1,element.numberOfExposures)
    count = det.GetFieldCount()
    endCount = count+element.numberOfExposures+1
    if(element.kVp != 0):
        cab.Configure(element.kVp,element.xOnTimeSecs*1000)
        if cab.FireXRay() != True:
            return -1
        cab.WaitXRayOn()
    
    det.SoftwareTrigger()
    print("Acquiring Image!")
    
    while (count<endCount) :
        count = det.GetFieldCount()
        print("."),
        time.sleep(0.2)
    
    cab.WaitXRayOff()
    print("Image Captured!")
    for i in range (0,element.numberOfExposures):
        det.ReadBuffer(i,img,i)
        
    img.UnscrambleImage()
    
    _kVp = "%.0f" % element.kVp    
    _t_expms = "%.2f" % element.t_expms
    
    if(element.comment != None):
        filename += "_"+str(_kVp)+"kVp_" + str(_t_expms) + "ms_300uA_" + element.comment + "_exp" + str(expNum).zfill(4)
    else:
        filename += "_"+str(_kVp)+"kVp_" + str(_t_expms) + "ms_300uA_exp" + str(expNum).zfill(4)
    
    wellOffset = 0
    if element.fullWell == DexelaPy.FullWellModes.Low:
        wellOffset =3
    if element.Binning == DexelaPy.bins.x11:
        binOffset = 0
    elif element.Binning == DexelaPy.bins.x22:
        binOffset = 1
    elif element.Binning == DexelaPy.bins.x44:
        binOffset = 2
    
    offset = wellOffset+binOffset
    
    if element.command == 'dark':
        img.FindMedianofPlanes()
        filename = outpath+"Darks\\" + filename + "_dark.smv"
        img.WriteImage(filename)    
        darkIms[offset] = filename
    elif element.command == 'flood':
        filename = outpath+"Floods\\" + filename + "_flood.smv"       
        if os.path.isfile(darkIms[offset]):    
            tmpImage = DexelaPy.DexImagePy()
            tmpImage.LoadDarkImage(darkIms[offset])
            tmpImage.LoadFloodImage(img)
            flood = tmpImage.GetFloodImage()
            flood.WriteImage(filename)    
            gainIms[offset] = filename
        else:
            print("Cannot capture flood without first capturing dark!")

        
           
        
def Linearity(outpath,det,element,expNum):
    _t_expms = "%.1f" % element.t_expms
    print("Capturing sequence with:")
    print("KV: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    print("Well-Mode: " + str(element.fullWell))
    print("Binning-Mode: " + str(element.Binning))
    print("Exposure-Time: " + str(element.t_expms))
    print("Number of Exposures: " + str(element.numberOfExposures))
    img = DexelaPy.DexImagePy()
    det.SetFullWellMode(element.fullWell)    
    det.SetBinningMode(element.Binning)
    det.SetExposureTime(element.t_expms)
    det.SetTriggerSource(element.trigger)
    det.SetNumOfExposures(element.numberOfExposures)
    
    det.GoLiveSeq(0,element.numberOfExposures-1,element.numberOfExposures)
    count = det.GetFieldCount()
    endCount = count+element.numberOfExposures+1
    det.SoftwareTrigger()
    print("Acquiring Image!")
    
    while (count<endCount) :
        count = det.GetFieldCount()
        print("."),
        time.sleep(0.2)
    
    print("Image Captured!")
    for i in range (0,element.numberOfExposures):
        det.ReadBuffer(i,img,i)
        
    img.UnscrambleImage()
    
    filename = outpath+"Linearity\linearity.txt"
    Utils.LinMeasurement(det,img,filename)
    
    
        
def RunScript(det,list,portNum):
    global cab
    
    print("Connecting to detector...")
    det.OpenBoard()
    det.SetExposureMode(DexelaPy.ExposureModes.Sequence_Exposure)
    model = det.GetModelNumber()
    serial = det.GetSerialNumber()
    filename = str(model)+"-"+str(serial)+"-"
    outpath = "Output\\"
    #outpath = "C:\\Users\\plagon\\Documents\\Projects\\test\\Array_1207_454\\X-Ray\\"
    
    counter = 1
    
    for element in list:
        if element.isCommand == False or element.command == 'dummy':
            TestSequence(outpath,filename,det,cab,element,counter)
            print("\n")
        elif element.command == 'dark'  or element.command == 'flood':
            DarkFloodCapture(outpath,filename,det,cab,element,counter)
        elif element.command == 'generatorinit':
            print("Initializing Generator!")
            cab = Faxitron.Cabinet(portNum)
        elif element.command == 'autocal':
            print("Calibrating ADC Offsets!")
            tempFileName = outpath+"AutoCal\\" + filename + "_autocal.txt"
            Utils.GetDarkOffsets(det,tempFileName)
            print("\n")
        elif element.command == 'wait':
            print("sleeping for: " + str(element.t_expms) + "ms")
            time.sleep(int(element.t_expms/1000))
            print("\n")
        elif element.command == 'alert':
            winsound.Beep(3000,2000)
            easygui.msgbox(element.comment,title="Alert!")
        elif element.command == 'message':
            easygui.msgbox(element.comment,title="Alert!")
        elif element.command == 'lightbox':
            lb = LightBox.LightBox(portNum)
            lb.SetIntensity(int(element.comment))
        elif element.command == 'lin-measurement':
            Linearity(outpath,det,element,counter)
        elif element.command == 'startreport':
            if os.path.isfile("XDComs.exe") and os.path.isfile("XDMessaging.dll"):
                exeCommand = "XDComs.exe " + "FromSCap " + outpath
                os.system(exeCommand)
            else:
                print("XDComs or dependency is missing!")
        elif element.command == 'loaddefectmaps':
            LoatDefectMaps(outpath)
        counter += 1
        
    if cab != None:
        cab.close()
    time.sleep(0.5)
    print("Image successfully saved!")
    det.CloseBoard()
    
    
    return 

try:
    global cab
    cab = None
    #list = DexList.DexList('linearity.xml').xmlList
    list = DexList.DexList('test-stage-210-65-report.xml').xmlList
    #list = DexList.DexList('test-stage-2-defect-map-autoload.xml').xmlList
    #list = DexList.DexList('singleexposure.xml').xmlList
    
    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    for i in range(0,count):
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        print("Running XML script with detector with serial number: %d" % info.serialNum)
        RunScript(det,list,3)

except DexelaPy.DexelaExceptionPy as ex:
    print("Exception Occurred!")
    print("Description: %s" % ex)
    DexException = ex.DexelaException
    global cab
    if cab != None:
        cab.close()
    print("Function: %s" % DexException.GetFunctionName())
    print("Line number: %d" % DexException.GetLineNumber())
    print("File name: %s" % DexException.GetFileName())
    print("Transport Message: %s" % DexException.GetTransportMessage())
    print("Transport Number: %d" % DexException.GetTransportError())
except Exception:
    print("Exception OCCURRED!")
    global cab
    if cab != None:
        cab.close()




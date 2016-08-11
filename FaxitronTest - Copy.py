import DexelaPy
import Utils
import DexList
import Faxitron
import time
import copy
import winsound
import easygui
import LightBox
from FileUtils import *
import os
import sys



global defectMaps
defectMaps = ["","",""]
global darkIms
darkIms = ["","","","","",""]
global gainIms
gainIms = ["","","","","",""]


dir_cfg ="C:\\CMOS\Configs\\"

dir_output_parent ="C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet"
dir_output =os.path.join(dir_output_parent,"Output")
dir_output_AutoCal =os.path.join(dir_output,"AutoCal")
dir_output_Darks =os.path.join(dir_output,"Darks")
dir_output_Floods =os.path.join(dir_output,"Floods")
dir_output_DefectMap =os.path.join(dir_output,"Defect Map")
dir_output_testRecord =os.path.join(dir_output,"Test Record")
dir_output_custDelivery =os.path.join(dir_output,"Customer Delivery")
dir_output_custDelivery_TestImages =os.path.join(dir_output_custDelivery,"Test Images")

drv_dvd ="E:"
dir_dataServer ="V:\\TestData\\ZT3\\CMOS\\data\\"


'''temp csv output file for processing result'''
temp_csv_output ="C:\\CMOSTestCSVOutput.csv";


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
    print("TestSequence: Capturing sequence with:")
    print("KV: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    print("Well-Mode: " + str(element.fullWell))
    print("Binning-Mode: " + str(element.Binning))
    print("Exposure-Time: " + str(element.t_expms))
    print("Number of Exposures: " + str(element.numberOfExposures))


    _kVp = "%.0f" % element.kVp    
    _t_expms = "%.2f" % element.t_expms

    ### sample name: 1512-15307-High-162.0ms-Binx11-GainCorrected--DefectCorrected-_45kVp_162.00ms_300uA_hand-phantom_exp0003.tif
    filename = filename + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning) 
    
    darkCorrect = element.darkCorrect
    if darkCorrect:
        print("Dark correction is enabled")
        #notInName- filename += "-DarkCorrected-"
    else:
        print("Dark correction is disabled")
    gainCorrect = element.gainCorrect
    if gainCorrect:
        print("Gain correction is enabled")
        filename += "-GainCorrected-"
    else:
        print("Gain correction is disabled")
    defectCorrect = element.defectCorrect
    if defectCorrect:
        print("Defect correction is enabled")
        filename += "-DefectCorrected-"
    else:
        print("Defect correction is disabled")

    #filename = outpath+filename + "_exp" + str(expNum).zfill(4) + ".tif"
    filename = filename + "_" +str(_kVp)+"kVp_300uA"
    if(element.comment != None):
        filename += "_" + element.comment 

    filename = outpath+filename + ".tif"
    print filename

    
    img = DexelaPy.DexImagePy()
    det.SetFullWellMode(element.fullWell)    
    det.SetBinningMode(element.Binning)
    det.SetExposureTime(element.t_expms)
    det.SetTriggerSource(element.trigger)
    det.SetNumOfExposures(element.numberOfExposures)
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
    
    if element.command != 'dummy':
        img.WriteImage(filename)
        
        
def dosimeas(cab,element):    
    print("dosi meas with:")
    print("KV: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    
    if(element.kVp != 0):
        cab.Configure(element.kVp,element.xOnTimeSecs*1000)
        if cab.FireXRay() != True:
            return -1
        cab.WaitXRayOn()

    time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
    cab.WaitXRayOff()

        
        
        
def DarkFloodCapture(outpath,filename,det,cab,element,expNum):
    global darkIms
    global gainIms
    
    _t_expms = "%.1f" % element.t_expms
    print("DarkFloodCapture: Capturing sequence with:")
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
    #filename = filename + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning)
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

    '''
    if(element.comment != None):
        #filename += "_"+str(_kVp)+"kVp_" + str(_t_expms) + "ms_300uA_" + element.comment + "_exp" + str(expNum).zfill(4)
        filename += "_"+str(_kVp)+"kVp_300uA_" + element.comment + "_exp" + str(expNum).zfill(4)
    else:
        #filename += "_"+str(_kVp)+"kVp_" + str(_t_expms) + "ms_300uA_exp" + str(expNum).zfill(4)
        filename += "_"+str(_kVp)+"kVp_300uA_exp" + str(expNum).zfill(4)
    '''
    
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
        #filename = outpath+"Darks\\" + filename + "_dark.smv"
        ### sample name: 1512-15307-High-162.0ms-Binx11-GainCorrected--DefectCorrected-_45kVp_162.00ms_300uA_hand-phantom_exp0003.tif
        filename =outpath+"Darks\\" + filename + str(element.fullWell)+"-"+str(element.numberOfExposures)+"x"+str(_t_expms)+"ms-Bin" +str(element.Binning)+ "_" + element.command +".smv"
        print filename        

        img.WriteImage(filename)    
        darkIms[offset] = filename
    elif element.command == 'flood':
        #filename = outpath+"Floods\\" + filename + "_flood.smv"
        filename =outpath+"Floods\\" + filename + str(element.fullWell)+"-"+str(element.numberOfExposures)+"x"+str(_t_expms)+"ms-Bin" +str(element.Binning)+"_" + element.command+".smv"
        print filename        
        
        if os.path.isfile(darkIms[offset]):    
            tmpImage = DexelaPy.DexImagePy()
            tmpImage.LoadDarkImage(darkIms[offset])
            tmpImage.LoadFloodImage(img)
            flood = tmpImage.GetFloodImage()
            flood.WriteImage(filename)    
            gainIms[offset] = filename
        else:
            print("Cannot capture flood without first capturing dark!")





def LagTest(outpath,filename,det,element,expNum):
    _t_expms = "%.1f" % element.t_expms
    print("LagTest: Capturing sequence with:")
    print("Well-Mode: " + str(element.fullWell))
    print("Binning-Mode: " + str(element.Binning))
    print("Exposure-Time: " + str(element.t_expms))
    print("Number of Exposures: " + str(element.numberOfExposures))

    starttime =time.time()

    _kVp = "%.0f" % element.kVp    
    _t_expms = "%.2f" % element.t_expms

    ### sample name: 1512-15307-High-162.0ms-Binx11-GainCorrected--DefectCorrected-_45kVp_162.00ms_300uA_hand-phantom_exp0003.tif
    filename = filename + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning) 
    
    darkCorrect = element.darkCorrect
    if darkCorrect:
        print("Dark correction is enabled")
        #notInName- filename += "-DarkCorrected-"
    else:
        print("Dark correction is disabled")
    gainCorrect = element.gainCorrect
    if gainCorrect:
        print("Gain correction is enabled")
        filename += "-GainCorrected-"
    else:
        print("Gain correction is disabled")
    defectCorrect = element.defectCorrect
    if defectCorrect:
        print("Defect correction is enabled")
        filename += "-DefectCorrected-"
    else:
        print("Defect correction is disabled")

    filename = filename + "_" +str(_kVp)+"kVp_300uA"
    if(element.comment != None):
        filename += "_" + element.comment 

    filename = outpath+filename 
    #print filename

    for imgIndex in range(0,9):
        print imgIndex
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
                
        print filename+"_" +str(imgIndex) +"_"+ str("%.0f" %(time.time() - starttime))+"sec.tif"
        img.WriteImage(filename+"_" +str(imgIndex) +"_"+ str("%.0f" %(time.time() - starttime) )+"sec.tif")
        time.sleep(1)

    return


            
           
        
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
        time.sleep
    img.UnscrambleImage()
    
    filename = outpath+"Linearity\linearity.txt"
    Utils.LinMeasurement(det,img,filename)
    
    
def getDetInfo():
    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    for i in range(0,count):
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        det.OpenBoard()
        model = det.GetModelNumber()
        serial = det.GetSerialNumber()
        det.CloseBoard()

        #print("Test detector with serial number: %d " % info.serialNum)
        print "Test detector with model:", model, "serial number:", serial
        return model, serial

    return 0, 0 #no detector found


def RunScriptDosimeter(list,portNum):
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
        else:
            print(element.command," is not support in this test")
        
    if cab != None:
        cab.close()
    time.sleep(0.5)
    print "All Done."
    
    return



def RunScript(det,list,portNum):
    global cab
    
    print("Connecting to detector...")
    det.OpenBoard()
    det.SetExposureMode(DexelaPy.ExposureModes.Sequence_Exposure)
    model = det.GetModelNumber()
    serial = det.GetSerialNumber()
    filename = str(model)+"-"+str(serial)+"-"
    outpath = "Output\\"
    
    counter = 1
    
    for element in list:
        if element.isCommand == False or element.command == 'dummy':
            TestSequence(outpath,filename,det,cab,element,counter)
            print("\n")
        elif element.command == 'dark'  or element.command == 'flood':
            DarkFloodCapture(outpath,filename,det,cab,element,counter)
        elif element.command == 'lag':
            LagTest(outpath,filename,det,element,counter)
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
        elif element.command == 'dosimeas':
            dosimeas(cab,element)
            print("\n")
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


def vfyDetectorSerialNum(serialNum):
    ### verify the detector serial number
    if not easygui.ynbox('Detector serial No.: '+str(serialNum)+'\nDoes it match with record?',
                         'Detector Serial No. Verify', ('Yes', 'No'), None):
        pkiWarningBox("Detector has wrong serial number. It failed test.")
        sys.exit(0)

def DetectorTest(cfgFile) :
    print "loading test sequence from ", cfgFile

    if not os.path.isfile(cfgFile):
        print cfgFile, "is not a file"
        
    list = DexList.DexList(cfgFile).xmlList
    
    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    for i in range(0,count):
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        print("Running XML script with detector with serial number: %d" % info.serialNum)
        RunScript(det,list,3)
    
    return

try:
    global cab
    cab = None
    
    cfgFile ='none'
    dir_cfg_model =dir_cfg

    ### get detector info
    foundDetector=False
    detInfo =getDetInfo()
    det_model =detInfo[0]
    det_sn =detInfo[1]
    if det_model==0 | det_sn==0:
        print "No detector found."
    else:
        print "connected detector mode:", det_model, "serial:", det_sn
        foundDetector=True
        dir_cfg_model =dir_cfg + str(det_model) +'\\'

    selection =easygui.buttonbox('Select Test Phase.', 'Test Phase',
                                 ('Phase 1', 'Phase 2', 'Data Transfer', 'Eng'))
    if selection =='Phase 1':
        if not foundDetector:
            pkiWarningBox("No detector found. Need Detector for this test. Test stopped!")
            sys.exit(0)
    
        vfyDetectorSerialNum(det_sn)

        '''check whehter linearity data is on the server and ready for report'''
        result =getLinearityData(dir_dataServer, dir_output_parent, str(det_model),str(det_sn))
        if not result[0]:
            easygui.msgbox("Linearity Data not found. Test failed and stop here.",
                           "Incoming File Check Error", 'OK', None, None)
            #sys.exit(0)

        cfgFile =dir_cfg_model + 'cmos_' + str(det_model)+'_seq1.xml'
        print "loading cfg file", cfgFile
        #clean up old files, if required.
        selection =easygui.buttonbox('Directory Clean-up', 'Delete old files?', ('Yes', 'No'))
        if selection =='Yes':
            prepDir(dir_output)
            prepDir(dir_output_AutoCal);
            prepDir(dir_output_Darks);
            prepDir(dir_output_Floods);
            prepDir(dir_output_DefectMap);
        print "start Detector Test..."
        DetectorTest(cfgFile)

    elif selection == 'Phase 2':
        if not foundDetector:
            easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
            sys.exit(0)

        vfyDetectorSerialNum(det_sn)
        
        cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq2.xml"

        DetectorTest(cfgFile)
        
    elif selection =="Data Transfer":
        if not foundDetector:
            easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
            sys.exit(0)

        #update CofC with linearity data
        bRelease =False
        result =updateCofCLinearityData(os.path.join(dir_output_parent, 'Linearity'), dir_output_testRecord, temp_csv_output)
        if result[0]:
            easygui.msgbox(result[1], "Final Test Result", 'OK', None, None)
            if result[1]=="PASS":
                bRelease=True
        else:
            pkiWarningBox("Failed update CofC. Test Failed")

        if bRelease:
            copyImages2Delivery(dir_output, dir_output_custDelivery_TestImages)     
            if not isDataReadyForDelivery(dir_output_custDelivery):
                print dir_output_custDelivery, " not found. Please check. Data Transfer Failed!"
                sys.exit(0)
        
        #copy data to server
        data2Server(dir_output, dir_dataServer, str(det_model), str(det_sn))





        '''
        if bRelease:
            easygui.msgbox("Insert a blank CD to attached CD drive and wait for the next instruction", "Alert", 'OK', None, None)
            time.sleep(60) #wait for 60s to allow CD 
            count=0
            while (not easygui.ynbox("Has CD ready yet?", "Alert") and count<50):
                print "wait for a second. Press 'Yes' to continue when it is ready";
                time.sleep(1)
                count +=1
                
            print "CD is ready and proceed to copy data to CD"
                          
            if not isDiskReady(drv_dvd, str(det_model)+'_'+str(det_sn)):
                print "CD is not ready. Please check"
                sys.exit(0)
            data2Dvd(dir_output_custDelivery, drv_dvd)
        '''
        
        easygui.msgbox("File Transfer Completed", "File Transfer Done", 'OK', None, None)
        
    elif selection =='Eng':
        phantomSel =easygui.buttonbox('Select Test.', 'Eng Tests',
                                 ('Hand', 'Digiman', 'CIRS', 'TOR-MAS','MTF','EXP', 'Dosimeter'))
        if phantomSel=="Hand":
            if not foundDetector:
                easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
                sys.exit(0)

            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq_phantom_hand.xml"

            DetectorTest(cfgFile)
        elif phantomSel=="Digiman":
            if not foundDetector:
                easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
                sys.exit(0)

            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq_phantom_Digiman.xml"

            DetectorTest(cfgFile)
        elif phantomSel=="CIRS":
            if not foundDetector:
                easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
                sys.exit(0)

            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq_phantom_CIRS.xml"

            DetectorTest(cfgFile)
        elif phantomSel=="TOR-MAS":
            if not foundDetector:
                easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
                sys.exit(0)

            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq_phantom_TOR_MAS.xml"

            DetectorTest(cfgFile)
        elif phantomSel=='MTF': 
            if not foundDetector:
                easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
                sys.exit(0)

            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq_phantom_MTF.xml"

            DetectorTest(cfgFile)
        elif phantomSel=='EXP': 
            if not foundDetector:
                easygui.msgbox("No detector found. Need Detector for this test. Test stopped!", "Warning...", 'OK', None, None)
                sys.exit(0)

            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq_exp.xml"

            DetectorTest(cfgFile)            
        else:
            cfgFile =dir_cfg + 'cal_seq_FaxitronXRay_dosimeter.xml'

            print "loading test sequence from ", cfgFile
            list = DexList.DexList(cfgFile).xmlList

            RunScriptDosimeter(list,3)

            sys.exit(0)
            
    else:
        print "Not support", selection
        #sys.exit(0)

except DexelaPy.DexelaExceptionPy as ex:
    #global cab
    print("Exception Occurred!")
    print("Description: %s" % ex)
    DexException = ex.DexelaException
    if cab != None:
        cab.close()
    print("Function: %s" % DexException.GetFunctionName())
    print("Line number: %d" % DexException.GetLineNumber())
    print("File name: %s" % DexException.GetFileName())
    print("Transport Message: %s" % DexException.GetTransportMessage())
    print("Transport Number: %d" % DexException.GetTransportError())
except Exception:
    print("Exception OCCURRED!")
    #global cab
    if cab != None:
        cab.close()




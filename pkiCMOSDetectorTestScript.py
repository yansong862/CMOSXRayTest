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




import DexelaPy 
import DexElement
from DexelaDetectorFunc import *

import Utils
import DexList
import Faxitron
import LightBox
from FileUtils import *
from pkiImageProc import *



global defectMaps
defectMaps = ["","",""]
global darkIms
darkIms = ["","","","","",""]
global gainIms
gainIms = ["","","","","",""]



###
### production/eng configurations
###
global gbReleaseCD, gbLoadDB
global giServer 

gbReleaseCD =True
gbLoadDB =True
giServer =0 #0-production, 1- engineer (ZT3), 2- Eng debug



###
### const
###

gDefaultCFGFile ='NA'
gDefaultVal_RMANo ='NA'

gDefaultVal_cfgFileName ='NA'
gDefaultVal_detModel =0
gDefaultVal_detSN =0
gDefaultVal_detSNStr =str(gDefaultVal_detSN)
gDefaultVal_buildLabel ='NA'
gDefaultVal_buildRev ='NA'
gDefaultVal_ADCBoardType ='NA'
gDefaultVal_ADCFirmRev ='NA'
gDefaultVal_DAQFirmRev ='NA'


gDetInfoIndex_found =0
gDetInfoIndex_model =1
gDetInfoIndex_serialNo =2
gDetInfoIndex_serialNoStr =3
gDetInfoIndex_rmaNo =4
gDetInfoIndex_buildLabel =5
gDetInfoIndex_buildRev=6
gDetInfoIndex_ADCBoardType =7
gDetInfoIndex_ADCFirmRev=8
gDetInfoIndex_DAQFirmRev=9
gDetInfoIndex_isGigE =10
gDetInfoIndex_isSlowMode =11
gDetInfoIndex_MTFResolution=12
gDetInfoIndex_GigENoisSpec=13




###
### working directories list
###
dir_cfg ="C:\\CMOS\Configs\\"
dir_cmos_report ="C:\\CMOS\\ForReports\\"
dir_xray_report_generator ="C:\\Users\\scltester\\Desktop\\XRayReportGenerator\\"

dir_temp_server ="\\\\optsclf01\\public\\CMOS\\XRay\\"


global dir_output_parent                        #="C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet"
global dir_output                               #=os.path.join(dir_output_parent,"Output")
global dir_output_AutoCal                       #=os.path.join(dir_output,"AutoCal")
global dir_output_Darks                         #=os.path.join(dir_output,"Darks")
global dir_output_Floods                        #=os.path.join(dir_output,"Floods")
global dir_output_DefectMap                     #=os.path.join(dir_output,"Defect Map")
global dir_output_testRecord                    #=os.path.join(dir_output,"Test Record")
global dir_output_custDelivery                  #=os.path.join(dir_output,"Customer Delivery")
global dir_output_custDelivery_TestImages       #=os.path.join(dir_output_custDelivery,"Test Images")
global dir_output_custDelivery_SupportFiles     #=os.path.join(dir_output_custDelivery,"Support Files")



drv_dvd ="E:"

dir_inDataServer ="\\\\amersclnas02\\Fire\\CMOS\\data\\"
dir_inRMADataServer ="\\\\amersclnas02\\Fire\\CMOS\\RMAs\\"
dir_UKCofCs ="\\\\optsclf01\\public\\ysong\\TestData\\UKCofC"
if giServer==0:
    dir_outDataServer ="\\\\amersclnas02\\Fire\\CMOS\\data\\"
    dir_outRMADataServer="\\\\amersclnas02\\Fire\\CMOS\\RMAs\\"
elif giServer==1:    
    dir_outDataServer ="V:\\TestData\\ZT3\\CMOS\\data\\"
    dir_outRMADataServer="V:\\TestData\\ZT3\\CMOS\\data\\"
else:
    dir_outDataServer ="\\\\optsclf01\\public\\ysong\\TestData\\"
    dir_outRMADataServer ="\\\\optsclf01\\public\\ysong\\TestData\\"

dir_linear ="C:\\Users\\scltester\\Documents\\GitWorkspace\\FaxitronCabinet\\Linearity"
dir_dataEngServer ="\\\\optsclf01\\public\\ysong\\TestData\\"

'''temp csv output file for processing result'''
temp_csv_output ="C:\\CMOSTestCSVOutput.csv";

XRayCabDosiLog ="c:\\cmosXRayCabDosiLog.txt"


_fdelay_getFieldCount =0.02

###
### functions
###

def getCurrentDir():
    global dir_output_parent
    global dir_output
    global dir_output_AutoCal
    global dir_output_Darks
    global dir_output_Floods
    global dir_output_DefectMap
    global dir_output_testRecord
    global dir_output_custDelivery
    global dir_output_custDelivery_TestImages
    global dir_output_custDelivery_SupportFiles
    
    dir_output_parent =os.getcwd() #os.chdir(os.path.dirname(os.getcwd()))
    "current dir:", dir_output_parent
    dir_output =os.path.join(dir_output_parent,"Output")
    dir_output_AutoCal =os.path.join(dir_output,"AutoCal")
    dir_output_Darks =os.path.join(dir_output,"Darks")
    dir_output_Floods =os.path.join(dir_output,"Floods")
    dir_output_DefectMap =os.path.join(dir_output,"Defect Map")
    dir_output_testRecord =os.path.join(dir_output,"Test Record")
    dir_output_custDelivery =os.path.join(dir_output,"Customer Delivery")
    dir_output_custDelivery_TestImages =os.path.join(dir_output_custDelivery,"Test Images")
    dir_output_custDelivery_SupportFiles =os.path.join(dir_output_custDelivery, "Support Files")
    return


def startReportApp(img_dir):
    theApp =dir_output_parent+"\\XDComs.exe"
    theDLL =dir_output_parent+"\\XDMessaging.dll"
    if os.path.isfile(theApp) and os.path.isfile(theDLL):
        exeCommand = theApp + " FromSCap " + img_dir
        os.system(exeCommand)
        return True
    else:
        print("XDComs or dependency is missing!")
        return False

    
    
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


def DefectCorrectByName(dir_image, namePattern):
    try:
        bDebug =False
        
        if not os.path.exists(dir_image):
            print "Cannnot find directory: ", dir_image
        print dir_image
        
        dir_defectmap =os.path.join(dir_image, "Defect Map")
        if not os.path.exists(dir_defectmap):
            print "Cannnot find directory: ", dir_defectmap
        print dir_defectmap

        dir_darks =os.path.join(dir_image, "Darks")
        if not os.path.exists(dir_darks):
            print "Cannnot find directory: ", dir_darks
        print dir_darks

        dir_floods =os.path.join(dir_image, "Floods")
        if not os.path.exists(dir_floods):
            print "Cannnot find directory: ", dir_floods
        print dir_floods
            
    
        defectFileNamePattern ="None"
        darksFileNamePattern ="None"
        floodsFileNamePattern ="None"
        for filename in glob.glob(os.path.join(dir_image,namePattern)):
            if bDebug: print filename
            if 'Binx11' in filename:
                defectFileNamePattern ="*Defect_Map_1x1.smv"
                if 'High' in filename:
                    darksFileNamePattern ="*High*Binx11*dark.smv"
                    floodsFileNamePattern ="*High*Binx11*flood*.smv"
                elif 'Low' in filename:
                    darksFileNamePattern ="*Low*Binx11*.smv"
                    floodsFileNamePattern ="*Low*Binx11*flood*.smv"
            elif 'Binx22' in filename:
                defectFileNamePattern ="*Defect_Map_2x2.smv"
                if 'High' in filename:
                    darksFileNamePattern ="*High*Binx22*dark.smv"
                    floodsFileNamePattern ="*High*Binx22*flood*.smv"
                elif 'Low' in filename:
                    darksFileNamePattern ="*Low*Binx22*.smv"
                    floodsFileNamePattern ="*Low*Binx22*flood*.smv"
            elif 'Binx44' in filename:
                defectFileNamePattern ="*Defect_Map_4x4.smv"
                if 'High' in filename:
                    darksFileNamePattern ="*High*Binx44*dark.smv"
                    floodsFileNamePattern ="*High*Binx44*flood*.smv"
                elif 'Low' in filename:
                    darksFileNamePattern ="*Low*Binx44*.smv"
                    floodsFileNamePattern ="*Low*Binx44*flood*.smv"                
            else:
                print "Un-expected file name. Test stopped.", filename
                return

            if defectFileNamePattern=="None" or \
               darksFileNamePattern=="None" or \
               floodsFileNamePattern=="None":
                print "one or more correction files not found."
                return

            
            if bDebug: print "FileNamePatterns:", darksFileNamePattern, floodsFileNamePattern, defectFileNamePattern

            defectMapName ="None"
            darksFileName ="None"
            floodssFileName ="None"
            
            defectMapCnt=0
            for item in glob.glob(os.path.join(dir_defectmap,defectFileNamePattern)):
                print item
                defectMapName =item
                defectMapCnt +=1
                                  
            if defectMapCnt!=1:
                print "There should be only one defect map found, but actual number is: ", defectMapCnt, " Test Stopped."
                return

            defectMapCnt=0
            for item in glob.glob(os.path.join(dir_darks,darksFileNamePattern)):
                if bDebug: print item
                darksFileName =item
                defectMapCnt +=1

            if defectMapCnt!=1:
                print "There should be only one dark map found, but actual number is: ", defectMapCnt, " Test Stopped."
                return

            defectMapCnt=0
            for item in glob.glob(os.path.join(dir_floods,floodsFileNamePattern)):
                if bDebug: print item
                floodsFileName =item
                defectMapCnt +=1
                                  
            if defectMapCnt!=1:
                print "There should be only one flood map found, but actual number is: ", defectMapCnt, " Test Stopped."
                return

            if defectMapName=="None" or \
               darksFileName=="None" or \
               floodsFileName=="None":
                print "one or more correction files not found."
                return
            

            print "CorrectionFileName:", darksFileName, floodsFileName, defectMapName
            
            
            print "Reading in input image: ", filename
            img =DexelaPy.DexImagePy(filename)
            #print img.GetImageType()
            #print img.GetImageBinning()
            
            print "Reading in dark image...", darksFileName
            img.LoadDarkImage(darksFileName)
            print("Subtracting dark from input...")
            img.SubtractDark()
            print("Dark Correction Success!")

            print "Reading in flood image...", floodsFileName
            img.LoadFloodImage(floodsFileName)
            print("Performing flood correction...")
            img.FloodCorrection()
            print("Gain Correction Success!")
            
            print "Reading in defect map: ", defectMapName
            img.LoadDefectMap(defectMapName)
            print("Performing defect correction...")
            img.DefectCorrection()
            print("Defect correct Success!")

            #print("Performing defect correction...")
            #img.FullCorrection()
            #print("Full correct Success!")
            
            print "Saving corrected image:", filename
            img.WriteImage(filename)
            print("Success!")        
            
            print ("\n")

            
        return

    except DexelaPy.DexelaExceptionPy as ex:
        print("Exception Occurred!")
        print("Description: %s" % ex)
        DexException = ex.DexelaException
        print("Function: %s" % DexException.GetFunctionName())
        return
    except Exception:
        print("Exception OCCURRED!")
        return



def DefectCorrect(dir_image):
    DefectCorrectByName(dir_image, "*DefectCorrected*.tif")
    return


def DefectCorrectSysNoiseImages(dir_image):
    DefectCorrectByName(dir_image, "*dark_1*.tif")
    DefectCorrectByName(dir_image, "*dark_2*.tif")
    return



def getXRaySetting():
    msg = "Enter X-Ray kVp and On-Time"
    title = "X-Ray Setting"
    fieldNames = ["\t kVp","\t On-Time"]
    #fieldValues = []  # we start with blanks for the values
    fieldValues = [50, 2]  # we start with defaults for the values
    fieldValues = easygui.multenterbox(msg,title, fieldNames,fieldValues)
    print "Entered: ", fieldValues

    # make sure that none of the fields was left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
            #print type(fieldValues[i])
            if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
            elif not fieldValues[i].isdigit():
                errmsg =errmsg + ('"%s" must be a number.\n\n' % fieldNames[i].strip())
                    
        if errmsg == "": break # no problems found
        fieldValues =easygui.multenterbox(errmsg,title, fieldNames,fieldValues)

    print "Entered: ", fieldValues
    return fieldValues


def goFireXRay(kVpSetting, XRayOnTime):
    print("Fire X-Ray for " + str(XRayOnTime)+ "sec at " + str(kVpSetting)+ " kVp")
    if kVpSetting ==0 or XRayOnTime ==0:
        print("Either kVp or X-Ray OnTime is '0'. No X-Ray fired")
        return False
    
    cab = Faxitron.Cabinet(3)

    cab.Configure(kVpSetting,XRayOnTime*1000)
    if cab.FireXRay() == True:
        cab.WaitXRayOn()

        time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
        cab.WaitXRayOff()

        if cab != None:
            cab.close()
            
    print "X-Ray Off"
    return True



def getOneImageAcqSetting():
    msg = "Enter Image Acquisition Setting"
    title = "Image Acquisition Setting"
    fieldNames = ["\t kVp","\t On-Time", "\t=========", "\t IsCommand(Y/N)", "\t Command", "\t Number Frames", "Binning Mode(11,22,44)", "\t Exposure Time(in ms)", \
                  "\t FullWell Mode(H/L)", "\t DarkCorrect(Y/N)", "\t GainCorrect(Y/N)", "\t DefectCorrect(Y/N)"]
    #fieldValues = []  # we start with blanks for the values
    fieldValues = [45, 2,"==============",\
                   'N','flood', \
                   1, 11, 162,'H','N','N','N'\
                   ] # we start with defaults for the values

    idigKVp =0
    idigXRayOnTime =1
    idigNumFrame =5
    idigBinMode =6
    idigExpTime =7

    idigUseCommand =3
    idigCommand =4
    idigWell =8
    idigDarkCorrect =9
    idigGainCorrect =10
    idigDefCorrect =11

    print (fieldValues[idigUseCommand])
    print (fieldValues[idigCommand])
    print (fieldValues[idigWell])
    print (fieldValues[idigDarkCorrect])
    print (fieldValues[idigGainCorrect])
    print (fieldValues[idigDefCorrect])
    

    
    fieldValues = easygui.multenterbox(msg,title, fieldNames,fieldValues)
    print "Entered: ", fieldValues

    # make sure that none of the fields was left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
            print type(fieldValues[i])
            if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
            else:
                if i==idigKVp or \
                   i==idigXRayOnTime or \
                   i==idigNumFrame or \
                   i==idigBinMode or \
                   i==idigExpTime:
                    if not fieldValues[i].isdigit():
                        errmsg =errmsg + ('"%s" must be a number.\n\n' % fieldNames[i].strip())
                elif i==idigWell:
                    if not(fieldValues[i].upper() =='H' or fieldValues[i] =='L'):
                        errmsg =errmsg + ('"%s" only can be \"H\" or \"L\".\n\n' % fieldNames[i].strip())
                elif i==idigUseCommand or \
                     i==idigDarkCorrect or \
                     i==idigGainCorrect or \
                     i==idigDefCorrect:
                    if not(fieldValues[i].upper() =='Y' or fieldValues[i] =='N'):
                        errmsg =errmsg + ('"%s" only can be \"Y\" or \"N\".\n\n' % fieldNames[i].strip())
                    
        if errmsg == "": break # no problems found
        fieldValues =easygui.multenterbox(errmsg,title, fieldNames,fieldValues)
    
    print "Entered: ", fieldValues
    return fieldValues


def vfyImageIntensity(det, img, bGreaterThan, target):
    bEnableIntensityWarning =False
    
    res =Utils.GetImageAverage(det, img)

    cmosImgIntensityLog ="c:\\cmosImgIntensity.txt"
    if (os.path.isfile(cmosImgIntensityLog)):
        f =open(cmosImgIntensityLog, 'a')
    else:
        f =open(cmosImgIntensityLog, 'w+')

    if bGreaterThan:
        f.write(str(det.GetModelNumber())+","+str(det.GetSerialNumber())+","+" dark"+","+str("%.0f" %res)+"\n")
    else:
        f.write(str(det.GetModelNumber())+","+str(det.GetSerialNumber())+","+" dark"+","+str("%.0f" %res)+"\n")
    f.close()

    if bEnableIntensityWarning:
        if bGreaterThan:
            if res <=target:
                pkiWarningBox("Expect image intensity >" +str("%.1f" % target)+ "Actual meas.=" + str("%0.1f" % res))
        else:
            if res >=target:
                pkiWarningBox("Expect image intensity <" +str("%.1f" % target)+ "Actual meas.=" + str("%0.1f" % res))
    
    return
    

def detAutoCal(ADCBoardType, det, filename, darkOffset):
    try:
        print("Calibrating ADC Offsets!")
        #tempFileName = outpath+"AutoCal\\" + filename + "_autocal.txt"
        #darkOffset = int(element.comment)
        StopDarkCalibration = 15000
        StartDarkCalibration = 14000

        #print "ADCBoardType", ADCBoardType
        if  ADCBoardType == 'ADC':
            StopDarkCalibration = 15000
            StartDarkCalibration = 14000
        elif ADCBoardType == 'ADCRP':
            StopDarkCalibration = 5100
            StartDarkCalibration = 4800
        else:
            pkiWarningBox("Not support ADC Board type: " + ADCBoardType + ". Failed.")
            return False

        print(ADCBoardType, StartDarkCalibration, StopDarkCalibration, darkOffset)
        Utils.GetDarkOffsets(det,filename,StartDarkCalibration,StopDarkCalibration,darkOffset)
        print("\n")

        return True
    except:
        pkiWarningBox("Autocal with " + ADCBoardType + " Failed.")
        return False



    
def TestSequence(outpath,filename,det,cab,element,expNum, bSaveAsSMV=False):
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

    if element.numberOfExposures<2:
        ### sample name: 1512-15307-High-162.0ms-Binx11-GainCorrected--DefectCorrected-_45kVp_162.00ms_300uA_hand-phantom_exp0003.tif
        filename = filename + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning)
    else:
        filename = filename + str(element.fullWell)+"-"+str(_t_expms)+"x"+str(element.numberOfExposures)+"ms-Bin"+str(element.Binning)
    
    
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

    if bSaveAsSMV:
        filename =outpath+filename +".smv"
    else:
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

    if(element.kVp != 0):
        time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
    cab.WaitXRayOff() 
    print("Image Captured!")
    for i in range (0,element.numberOfExposures):
        det.ReadBuffer(i,img,i)
        
    img.UnscrambleImage()

    if ("flood" in filename) or \
        ("Flood" in filename) or \
        ("phantom" in filename):
        vfyImageIntensity(det, img, True, 1000)
    else:
        vfyImageIntensity(det, img, False, 1000)
    
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

    return filename


        
def dosimeas(cab,element):    
    print("dosi meas with:")
    print("KV: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    
    if(element.kVp != 0):
        cab.Configure(element.kVp,element.xOnTimeSecs*1000)
        if cab.FireXRay() != True:
            return -1
        cab.WaitXRayOn()

    if(element.kVp != 0):
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
        time.sleep(_fdelay_getFieldCount)

    if(element.kVp != 0):
        time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
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

    if ("flood" in filename) or \
        ("Flood" in filename) or \
        ("phantom" in filename):
        vfyImageIntensity(det, img, True, 1000)
    else:
        vfyImageIntensity(det, img, False, 1000)



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
            time.sleep(_fdelay_getFieldCount)
        
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


def kVpExpScanTest(outpath,filenameBase,det,cab,element,expNum, iScanOption):
    global darkIms
    global gainIms

    print str(element.fullWell)
    print str(element.t_expms)
    print str(element.kVp)
    
    resFileName =outpath+str(det.GetModelNumber())+"_"+\
                  str(det.GetSerialNumber())+"_RegionAverage_"+str(element.fullWell)
    print "recording file name: ", resFileName;
    if iScanOption==0:
        resFileName =resFileName+ "_kVpScan_" + str(element.t_expms)+".txt"
    elif iScanOption==1:
        resFileName =resFileName+ "_expScan" + str(element.kVp)+".txt"
    else:
        resFileName =resFileName+ "_kVpExpScan" + str(element.kVp)+".txt"

    print "recording file name: ", resFileName;
        
    if(os.path.isfile(resFileName)):
        os.remove(resFileName)
    
    _t_expms = "%.1f" % element.t_expms
    print("kVpExpScanTest: Capturing sequence with:")
    print("max KVp: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    print("Well-Mode: " + str(element.fullWell))
    print("Binning-Mode: " + str(element.Binning))
    print("Exposure-Time: " + str(element.t_expms))
    print("Number of Exposures: " + str(element.numberOfExposures))


    if iScanOption==0:
        kVpMin =10.0
        kVpMax =element.kVp
        kVpStep =5.0

        expmsMin =element.t_expms
        expmsMax =element.t_expms
        expmsStep =10.0
    elif iScanOption==1 :#expScan scan
        kVpMin =element.kVp
        kVpMax =element.kVp
        kVpStep =5.0

        expmsMin =37.0
        expmsMax =element.t_expms
        expmsStep =10.0
    else: #scan both kVp and expms
        kVpMin =10.0
        kVpMax =element.kVp
        kVpStep =5.0

        expmsMin =37.0
        expmsMax =element.t_expms
        expmsStep =10.0


    print "kVp range", kVpMin, "~", kVpMax, ", Step =", kVpStep
    print "expms range", expmsMin, "~", expmsMax, ", Step =", expmsStep


    theKVp =kVpMin
    while theKVp<=kVpMax:

        theExpms =expmsMin
        while theExpms<=expmsMax:
            _kVp ="%.0f" % theKVp
            _t_expms ="%.2f" % theExpms
            print "_kVp =", _kVp, ", _t_expms=", _t_expms

            filename = filenameBase + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning) 
            
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

            if(theKVp != 0):
                
                cab.Configure(theKVp,element.xOnTimeSecs*1000)
                if cab.FireXRay() != True:
                    return -1
                cab.WaitXRayOn()
        
            det.SoftwareTrigger()
            print("Acquiring Image!")
        
            while (count<endCount) :
                count = det.GetFieldCount()
                print("."),
                time.sleep(_fdelay_getFieldCount)
                
            if(theKVp != 0):
                time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
            cab.WaitXRayOff()
            print("Image Captured!")
            for i in range (0,element.numberOfExposures):
                det.ReadBuffer(i,img,i)
            
            img.UnscrambleImage()

            res =Utils.GetRegionAverage(det, img)
            if(os.path.isfile(resFileName)):
                f = open(resFileName,'a')
            else:
                f = open(resFileName,'w+')

            f.write(str(element.fullWell) +", " + str("%.2f" % theKVp) +", " +str("%.2f" % theExpms) +", ")
            for index in range(len(res)):
                #f.write(str(res[index]))
                f.write(str("%.2f" % res[index])+", ")
            f.write("\n")
            f.close()
                
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
                    
            img.WriteImage(filename+".tif")
            time.sleep(10)

            theExpms +=expmsStep

        theKVp +=kVpStep

    return


           
        



def floodDarkAltTest(outpath,filenameBase,det,cab,element,expNum):
    global darkIms
    global gainIms

    # element setting is for flood. Dark image will use default setting in this routine
    #  for now, dark image use the same detector setting with no x-ray on
    print str(element.fullWell)
    print str(element.t_expms)
    print str(element.kVp)

    now = time.ctime()
    parsed = time.strptime(now)
    #print time.strftime("%Y_%b_%d_%H_%M_%S", parsed)
    
    resFileName =outpath+str(det.GetModelNumber())+"_"+\
                  str(det.GetSerialNumber())+"_" + time.strftime("%Y_%b_%d_%H_%M_%S", parsed)+".csv"
    print "recording file name: ", resFileName;
        
    print("floodDarkAltTest: Flood Capturing sequence with:")
    print("kVp: " + str(element.kVp))  
    print("X-Ray Time: " + str(element.xOnTimeSecs))  
    print("Well-Mode: " + str(element.fullWell))
    print("Binning-Mode: " + str(element.Binning))
    print("Exposure-Time: " + str(element.t_expms))
    print("Number of Exposures: " + str(element.numberOfExposures))

    frameChar ="flood"
    binning =DexelaPy.bins.x11
    theKVp =0
    imgPairCnt =0
    imgPairCntMax =int(element.comment)
    print imgPairCntMax, " pair of flood/dark image will be captured."
    while imgPairCnt<imgPairCntMax:
        imgPairCnt +=1

        for ii in range(0,3):
            #capture flood/dark image
            #  for now, dark image use the same detector setting with no x-ray on
            _t_expms = "%.1f" % element.t_expms
            filename = filenameBase + str(element.fullWell)+"-"+str(_t_expms)+"ms-Bin"+str(element.Binning) 
                
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

            if ii==0:
                theKVp =element.kVp
                binning =DexelaPy.bins.x44
                frameChar ="flood"
            elif ii==1:
                theKVp =element.kVp
                binning =DexelaPy.bins.x11
                frameChar ="dummy"
            else:
                theKVp =0
                binning =DexelaPy.bins.x11
                frameChar ="dark"


            filename =filename + "_" +str(theKVp)+"kVp_300uA_" + frameChar + "_" + str(imgPairCnt)
            filename = outpath+filename 
            print filename        

            img = DexelaPy.DexImagePy()
            det.SetFullWellMode(element.fullWell)    
            #det.SetBinningMode(element.Binning)
            det.SetBinningMode(binning)
            det.SetExposureTime(element.t_expms)
            det.SetTriggerSource(element.trigger)
            det.SetNumOfExposures(element.numberOfExposures)
            det.GoLiveSeq(0,element.numberOfExposures-1,element.numberOfExposures)
            count = det.GetFieldCount()
            endCount = count+element.numberOfExposures+1

            if(theKVp != 0):                    
                cab.Configure(theKVp,element.xOnTimeSecs*1000)
                if cab.FireXRay() != True:
                    return -1
                cab.WaitXRayOn()
            
            det.SoftwareTrigger()
            print("Acquiring Image!")
            
            while (count<endCount) :
                count = det.GetFieldCount()
                print("."),
                time.sleep(_fdelay_getFieldCount)

            if(theKVp != 0):
                time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
            cab.WaitXRayOff()
            print("Image Captured!")
            for i in range (0,element.numberOfExposures):
                det.ReadBuffer(i,img,i)
                
            img.UnscrambleImage()

            res =Utils.GetImageAverage(det, img)
            if(os.path.isfile(resFileName)):
                f = open(resFileName,'a')
            else:
                f = open(resFileName,'w+')

            f.write(str(det.GetModelNumber())+"_" + str(det.GetSerialNumber())+", "+\
                        frameChar+ "," + str(element.fullWell) +", " + str("%.0f" % theKVp) +", " +str("%.0f" % element.t_expms) +", " +str("%.1f" % res))
                
            f.write("\n")
            f.close()
                    
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
                        
            img.WriteImage(filename+".tif")
            #time.sleep(1)

    return


           
def printImageInfo(image):
        print "type: ", image.GetImageType()
        print "Bin: ", image.GetImageBinning()
        print "Plane Ave:", image.FindAverageofPlanes()
        print "Plane Median:", image.FindMedianofPlanes()
        print "Image Depth:", image.GetImageDepth()
        print "Image Model:", image.GetImageModel()
        print "Image Pixel Type:", image.GetImagePixelType()
        print "Image X:", image.GetImageXdim()
        print "Image Y:", image.GetImageYdim()
        print "Is Average", image.IsAveraged()
        print "Is DarkCorrected", image.IsDarkCorrected()
        print "Is FloodCorrected", image.IsFloodCorrected()
        print "Is DefectCorrected", image.IsDefectCorrected()
        print "Is Empty", image.IsEmpty()
        print "Is Linearized:", image.IsLinearized()
        print "Plane 0 Avg: ", str("%0.1f" %image.PlaneAvg(0))
        
        

def printImgBlockAvg():
    global lastKvp;
    
    #open directory
    imgdir =str(easygui.diropenbox("Open Image Folder", "Open Folder"))

    namepattern =["*High*flood*", "*Low*flood*", "*High*dark*", "*Low*dark*"]

    f = open("c:\\tempOutput.csv",'w+')
    
    #get all *.tif files
    for i in range(0, len(namepattern)):
        for filename in glob.glob(os.path.join(imgdir,namepattern[i])):
            if os.path.isdir(filename):
                continue
            
            print "\n", filename        
            words = path_leaf(filename).split("-")
            exptime =words[3].split(".")
            image =DexelaPy.DexImagePy(filename)
            printImageInfo(image)

            if filename.find("GainCorrected")>=0:
                f.write("GainCorrected,"+ str(image.GetImageModel())+"," + str(image.GetImageBinning())+"," + namepattern[i]+", "+\
                                    str(exptime[0]) +", " +str("%.0f" % image.PlaneAvg(0))+"\n" )
            else:                
                f.write("NoGainCorrected,"+ str(image.GetImageModel())+"," + str(image.GetImageBinning())+"," + namepattern[i]+", "+\
                                    str(exptime[0]) +", " +str("%.0f" % image.PlaneAvg(0))+"\n" )
        
    f.close()

    getStatusOutput("notepad.exe c:\\tempOUtput.csv")
    
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
        time.sleep(_fdelay_getFieldCount)
    
    print("Image Captured!")
    for i in range (0,element.numberOfExposures):
        det.ReadBuffer(i,img,i)
        time.sleep
    img.UnscrambleImage()
    
    filename = outpath+"Linearity\linearity.txt"
    Utils.LinMeasurement(det,img,filename)
    


def RunScript(ADCBoardType, det,list,portNum):
    global cab
    
    print("Connecting to detector...")
    det.OpenBoard()
    det.SetExposureMode(DexelaPy.ExposureModes.Sequence_Exposure)
    model = det.GetModelNumber()
    serial = det.GetSerialNumber()

    det_sn_str =str(serial)
    if len(det_sn_str)==4:
        det_sn_str ="0"+str(serial)
        print "added leading 0 to serial number ", det_sn_str

    
    filename = str(model)+"-"+det_sn_str+"-"
    outpath = "Output\\"
    
    counter = 1
    
    for element in list:
        print "### command field", element.isCommand, element.command
        if element.isCommand == False or element.command == 'dummy':
            TestSequence(outpath,filename,det,cab,element,counter)
            print("\n")
        elif element.command == 'dark'  or element.command == 'flood':
            DarkFloodCapture(outpath,filename,det,cab,element,counter)
        elif element.command == 'lag':
            LagTest(outpath,filename,det,element,counter)
        elif element.command == 'kVPScan':
            kVpExpScanTest(outpath,filename,det,cab,element,counter, 0)
        elif element.command == 'expScan':
            kVpExpScanTest(outpath,filename,det,cab,element,counter, 1)
        elif element.command == 'kVpExpScan':
            kVpExpScanTest(outpath,filename,det,cab,element,counter, 2)
        elif element.command == 'FloodDarkAlt':
            floodDarkAltTest(outpath,filename,det,cab,element,counter)
        elif element.command == 'generatorinit':
            print("Initializing Generator!")
            cab = Faxitron.Cabinet(portNum)
        elif element.command == 'autocal':
            tempFileName = outpath+"AutoCal\\" + filename + "_autocal.txt"
            if not detAutoCal(ADCBoardType, det, tempFileName, int(element.comment)):
                break;
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
            startReportApp(dir_output)

        elif element.command == 'loaddefectmaps':
            LoatDefectMaps(outpath)
        counter += 1
        
    if cab != None:
        cab.close()
    time.sleep(0.5)
    print("End RunScript!")
    det.CloseBoard()
    
    
    return 


def RunScript_OnePhantom(det,list,portNum, thePhantom):
    global cab

    easygui.msgbox("Please place phantom for "+thePhantom+" Test. \nPress OK when it is ready.", "Info", 'OK', None, None)
    
    print("Connecting to detector...")
    det.OpenBoard()
    det.SetExposureMode(DexelaPy.ExposureModes.Sequence_Exposure)
    model = det.GetModelNumber()
    serial = det.GetSerialNumber()

    det_sn_str =str(serial)
    if len(det_sn_str)==4:
        det_sn_str ="0"+str(serial)
        print "added leading 0 to serial number ", det_sn_str

    
    filename = str(model)+"-"+det_sn_str+"-"
    outpath = "Output\\"
    
    counter = 1
    
    for element in list:
        print "### command field", element.isCommand, element.command
        if element.command == 'dummy':
            TestSequence(outpath,filename,det,cab,element,counter)
            print("\n")
        elif element.isCommand == False and element.comment ==thePhantom:
            TestSequence(outpath,filename,det,cab,element,counter)
            print("\n")
        elif element.command == 'generatorinit':
            print("Initializing Generator!")
            cab = Faxitron.Cabinet(portNum)
        elif element.command == 'wait':
            print("sleeping for: " + str(element.t_expms) + "ms")
            time.sleep(int(element.t_expms/1000))
            print("\n")
        elif element.command == 'message':
            easygui.msgbox(element.comment,title="Alert!")
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
    print("End RunScript!")
    det.CloseBoard()
    
    
    return


def RunOneTest(element, bSMVFile=True, bShowImage=True):
    global cab

    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    for i in range(0,count):
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        
        print("Connecting to detector...")
        det.OpenBoard()
        det.SetExposureMode(DexelaPy.ExposureModes.Sequence_Exposure)
        model = det.GetModelNumber()
        serial = det.GetSerialNumber()

        det_sn_str =str(serial)
        if len(det_sn_str)==4:
            det_sn_str ="0"+str(serial)
            print "added leading 0 to serial number ", det_sn_str

        
        filename = str(model)+"-"+det_sn_str+"-"
        outpath = "Output\\"

        if not os.path.exists(outpath):
            os.mkdir(outpath)
        
        counter = 1
        
        print("Initializing Generator!")
        cab = Faxitron.Cabinet(3)

        print ("### isCommand:", element.isCommand, "command =",element.command)
        imagefile =TestSequence(outpath,filename,det,cab,element,counter, bSMVFile)
        print "saved as ", imagefile
        if bShowImage:
            getStatusOutput("xia "+imagefile)
        
        if cab != None:
            cab.close()
            
        time.sleep(0.5)
        print("Done!\n")
        
        det.CloseBoard()    
    
    return


def RunReadTempInC():
    return 25.01
    global cab

    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    for i in range(0,count):
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        
        print("Connecting to detector...")
        det.OpenBoard()
        det.SetExposureMode(DexelaPy.ExposureModes.Sequence_Exposure)
        model = det.GetModelNumber()
        serial = det.GetSerialNumber()


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
        print("\n\nTemperatre:")
        print (str("%f" %measedT))
        det.WriteRegister(116,0) #enable temperature read

        
        det.CloseBoard()    
    
    return measedT


def captureSingleDummy():
    bRunExe=True
    singleImgPyScript ="C:\\Users\\Public\\Documents\\PerkinElmer\\PerkinElmer\\DexelaDetectorAPI\\Examples\\DexelaDetectorEx_Py\\SingleImageEx_Py.py"

    if bRunExe:
        singleImgPyScript =dir_output_parent+"\\bin\\SingleImageEx.exe"
        #theCmd = "echo Y\n | "+ singleImgPyScript
        theCmd = singleImgPyScript 
    else:
        singleImgPyScript =dir_output_parent+"\\bin\\SingleImageEx_Py.py"
        theCmd = "echo Y\n | python "+ singleImgPyScript

    if os.path.exists(singleImgPyScript):
        print("run command " + theCmd)
        #result = getStatusOutput(theCmd)
        os.system(theCmd)
    else:
        print("\n\nNot foudn"+singleImgPyScript)

    return


    
def getMasterDetInfo():
    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    #onlyNeedMasterInfo_NoNeedLoop- for i in range(0,count):
    if count>0:
        i=0 #master device
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        det.OpenBoard()

        ### detector info
        model = det.GetModelNumber()
        serial = det.GetSerialNumber()
        firmVer =det.GetFirmwareVersion()
        #NoNeedThisNow- firmBuild =det.GetFirmwareBuild(); print ("FirmwareBuild: ", firmBuild)
        ### it can read register 1 to 4 for model 1512??
        print "reg 1 read", det.ReadRegister(i, 1)
        print "reg 2 read", det.ReadRegister(i, 2)
        print "reg 3 read", det.ReadRegister(i, 3)
        print "reg 4 read", det.ReadRegister(i, 4)

        serial_str =str(serial)
        if len(serial_str)==4:
            serial_str ="0"+str(serial)
            print "added leading 0 to serial number ", serial_str        
        
        det.CloseBoard()

        #print("Test detector with serial number: %d " % info.serialNum)
        print "Test detector with model:", model, "serial number:", serial, "Firmware Version:", firmVer
        return model, serial, serial_str, firmVer

    return 0, 0, '', 0 #no detector found


### get detector: bFoundDetector, model, serialNo, serialNoStr, ADCFirmwareRev,
###               buildLabel, ADCBoardType, bGigE, bSlowMode
def getDetTestInfo():
    bFoundDetector =False
    iModel =gDefaultVal_detModel
    iSerialNo =gDefaultVal_detSN
    sSerialNo =gDefaultVal_detSNStr
    rmaNo =gDefaultVal_RMANo
    sBuildLabel =gDefaultVal_buildLabel
    sBuildRev =gDefaultVal_buildRev
    sADCBoardType =gDefaultVal_ADCBoardType
    iADCFirmRev =0      # read from detector
    sADCFirmRev=gDefaultVal_DAQFirmRev    # record in file
    sDAQFirmRev=gDefaultVal_ADCFirmRev
    bGigE =False
    bSlowMode =False
    bMTFHighResolution=False
    bNewGigENoiseSpec =False
    cfgDir ='None'
    
    detInfo =getMasterDetInfo()
    iModel =detInfo[0]
    iSerialNo =detInfo[1]
    sSerialNo =detInfo[2]
    iADCFirmRev =detInfo[3]
           
    if iModel==0 | det_sn==0:
        bFoundDetector =False
        print ("No detector found.")
    else:
        print ("connected detector mode:", iModel, "serial:", sSerialNo)
        Utils.setGlobals(iModel)
        bFoundDetector=True
        buildInfo =getModelADCBoardInfo(str(iModel), dir_cmos_report, dir_xray_report_generator)
        sBuildLabel =buildInfo[0]
        sBuildRev =buildInfo[1]
        sADCBoardType =buildInfo[2]
        sADCFirmRev =buildInfo[3]
        sDAQFirmRev=buildInfo[4]
        sGigE =buildInfo[5]
        sSlowMode =buildInfo[6]
        sMTFResolution =buildInfo[7]
        sTheHighNoiseSpec=buildInfo[8]

        if "Y" in sSlowMode:
            bSlowMode =True

        if "Y" in sGigE:
            bGigE =True

        if "H" in sMTFResolution:
            bMTFHighResolution=True
            
        if "Y" in sTheHighNoiseSpec:
            bNewGigENoiseSpec =True


    if bFoundDetector:
        ### verify master ADC firmware rev programmed in detector vs. record in file:
        thelist =re.split(r'[;,/a-zA-Z\s]\s*', sADCFirmRev)
        themasterADCRev=None
        for item in thelist:
            if len(item)>0:
                themasterADCRev =item
                break    

        if len(str(iADCFirmRev))<5: #Some units built in London has 5-digit ADCFirmware rev. Don't check these units.
            if not (str(iADCFirmRev) == themasterADCRev):
                if easygui.ynbox('Read ADC firmware rev '+str(iADCFirmRev)+' not match with record '+themasterADCRev+'.\n'+\
                                     'Continue test?\n(For production (including RMA), select <No> to stop)',\
                                     'Detector master ADC firmware rev Verify', ('Yes', 'No'), None):
                    ### ask user to key in teh firmware rev
                    sADCFirmRev =easygui.enterbox("Enter ADC Firmware Revision", "ADC Firmware Rev", '', True)
                else:
                    bFoundDetector =False

    if bFoundDetector: #detector is connected and the info is correct.
        #if easygui.ynbox("Is this RMA unit?", "Linearity", ('Yes', 'No'), None):
        if not easygui.ynbox("New or RMA", "Is RMA", ('New Build', 'RMA'), None):
            rmaNo =getRMANo()                    
            
    ### Note: if the order of return values change, please change definition of 'gDetInfoIndex_.....' constants
    return bFoundDetector, iModel, iSerialNo, sSerialNo, rmaNo,  sBuildLabel, sBuildRev, sADCBoardType, sADCFirmRev, sDAQFirmRev, bGigE, bSlowMode, bMTFHighResolution, bNewGigENoiseSpec            
    
    

def getOfflineDetInfo():
    bFoundDetector =False
    iModel =0
    iSerialNo =0
    sSerialNo =''
    sBuildLabel ='None'
    sBuildRev =None
    iADCFirmRev =0      # read from detector
    sADCFirmRev=None    # record in file
    sDAQFirmRev=None
    sADCBoardType ='None'
    bGigE =False
    bSlowMode =False
    bMTFHighResolution=False
    bNewGigENoiseSpec =False
    
           
    bFoundDetector=True
    buildInfo =getADCBoardInfo(dir_cmos_report, dir_xray_report_generator, False)
    sBuildLabel =buildInfo[0]
    print sBuildLabel
    sBuildRev =buildInfo[1]
    sADCBoardType =buildInfo[2]
    sADCFirmRev =buildInfo[3]
    sDAQFirmRev=buildInfo[4]
    sGigE =buildInfo[5]
    sSlowMode =buildInfo[6]
    sMTFResolution =buildInfo[7]
    sTheHighNoiseSpec=buildInfo[8]

    if "Y" in sSlowMode:
        bSlowMode =True

    if "Y" in sGigE:
        bGigE =True

    if "H" in sMTFResolution:
        bMTFHighResolution=True
            
    if "Y" in sTheHighNoiseSpec:
        bNewGigENoiseSpec =True


    ### get model number
    print sBuildLabel
    thelist =re.split(r'[-;,/a-zA-Z\s]\s*', sBuildLabel)
    print thelist
    iModel = int(thelist[0])
    print iModel
    Utils.setGlobals(iModel)

    ### get serial number
    iSerialNo =easygui.integerbox('Enter Serial No.', 'Detector Serial No.', '', 0, 999999999)
    sSerialNo =str(iSerialNo)
    if len(sSerialNo)==4:
        sSerialNo ="0"+str(iSerialNo)
        print("added leading 0 to serial number " + sSerialNo)       

    ### get RMA
    rmaNo =getRMANo()

    ### Note: if the order of return values change, please change definition of 'gOfflineDetInfoIndex_.....' constants
    return bFoundDetector, iModel, iSerialNo, sSerialNo, rmaNo, sBuildLabel, sBuildRev, sADCBoardType, sADCFirmRev, sDAQFirmRev, bGigE, bSlowMode, bMTFHighResolution, bNewGigENoiseSpec

def enterDetSerialNo():
    ### get serial number
    bGotSerialNo =False

    while not bGotSerialNo:
        iSerialNo =easygui.integerbox('Enter Serial No.', 'Detector Serial No.', '', 0, 999999999)
        bGotSerialNo =pkiYNBox('Is this Correct serial No.:' +str(iSerialNo), 'Verify Serial No.')
        
    sSerialNo =str(iSerialNo)
    if len(sSerialNo)==4:
        sSerialNo ="0"+str(iSerialNo)
        print("added leading 0 to serial number " + sSerialNo)

    return iSerialNo, sSerialNo


def selectModelNo():
    return easygui.choicebox("Model List", "Select a Model", ("1207","1512","2307","2315","2923"))

def getOfflineDet_model_serialNo():

    bFoundDetector =False
    sModel =None
    iSerialNo =None
    sSerialNo =None
    
    ### get model number
    sModel = selectModelNo()
    if sModel ==None:
        return bFoundDetector, sModel, iSerialNo, sSerialNo

    Utils.setGlobals(int(sModel))
    print('Model No:' + sModel)

    (iSerialNo, sSerialNo)=enterDetSerialNo()

    bFoundDetector =True
    
            
    return bFoundDetector, sModel, iSerialNo, sSerialNo


    
def ConfirmDetectorSerialNum(serialNum):
    ### verify the detector serial number
    return easygui.ynbox('Detector serial No.: '+str(serialNum)+'\nDoes it match with record?',
                         'Detector Serial No. Verify', ('Yes', 'No'), None)


def isDefectorConnected():
    detInfo =getMasterDetInfo()

    if detInfo[0] ==0 or detInfo[1]==0: #none of these values should be '0'
        pkiWarningBox("No detector found. Check the detect connection.")
        return False
    else:
        return True
    
    
    
def isSameDetector(detModel, detSN):
    detInfo =getMasterDetInfo()
    
    if detModel ==detInfo[0] and detSN ==detInfo[1]:
        print("same detector.")
        return True
    else:        
        print("different detector.\nModel:" + str(detModel) + "/"+str(detInfo[0])+ "\nser. No:" +str(detSN)+"/" +str(detInfo[1]))
        return False
    


    

def DetectorTest(cfgFile, ADCBoardType, onePhantom=None) :
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
        if onePhantom is None:
            RunScript(ADCBoardType, det,list,3)
        else: 
            RunScript_OnePhantom(det,list,3, onePhantom)
    return



def getCFGFileName(phase, det_model, bGigE, bSlowMode, buildLabel):
    dir_cfg_model =dir_cfg + str(det_model) +'\\'
    cfgFile =gDefaultCFGFile

    if 'Phase 1' in phase or \
       'Init Test' in phase :
        if buildLabel is not None and \
           "2923M-C13-HECC-2" in buildLabel:
            cfgFile =dir_cfg_model + 'cmos_2923M_C13_HECC_2_seq1'
        else:
            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq1"

    elif 'Phase 2' in phase:
        if buildLabel is not None and \
           "1512-C-3D-C600B" in buildLabel:
            cfgFile =dir_cfg_model + "cmos_" + buildLabel+"_seq2"
        elif buildLabel is not None and \
             "2923M-C13-HECC-2" in buildLabel:
            cfgFile =dir_cfg_model + 'cmos_2923M_C13_HECC_2_seq2'                 
        else:
            cfgFile =dir_cfg_model + "cmos_" + str(det_model)+"_seq2"
            
        
    if bGigE:
        cfgFile +="_GigE"

    if bSlowMode:
        cfgFile +="_Bin1x1"

    cfgFile +=".xml"

    if os.path.exists(cfgFile):
        print("use CFG file: ", cfgFile)
        return cfgFile
    else:
        pkiWarningBox(cfgFile+" not found. Cannot perform this test")
        return gDefaultCFGFile
        

def getRMANo():
    rmaNo =str(easygui.enterbox('Enter RMA No.', 'RMA No.', gDefaultVal_RMANo, True))
    while not (easygui.ynbox("Is this correct RMA No? " +rmaNo , "RMA No. Confirmation", ('Yes', 'No'), None)):
        rmaNo =str(easygui.enterbox('Enter RMA No.', 'RMA No.', rmaNo, True))
        if len(rmaNo) ==0:
            rmaNo =gDefaultVal_RMANo 

    print("RMA No.: " + rmaNo)
    return rmaNo

    
def isRMA(sRMANo):
    if sRMANo ==gDefaultVal_RMANo or \
       sRMANo is None:
        return False
    else:
        return True

    
    
def runPhase1Test(phase, det_model, det_sn_str, ramNo, ADCBoardType, bGigE, bSlowMode, buildLabel):
    #if easygui.ynbox("Delete old files?", "Directory Clean-up", ('Yes', 'No'), None):
    if True:    #skip the folder clean-up confirmation.
        prepDir(dir_output)
        prepDir(dir_output_AutoCal);
        prepDir(dir_output_Darks);
        prepDir(dir_output_Floods);
        prepDir(dir_output_DefectMap);


    if phase == "Phase 1":
        print ("Alignment check...")
        #el = DexElement.DexElement(False, '',0,1,DexelaPy.bins.x11,10038,DexelaPy.ExposureTriggerSource.Internal_Software,0,300,'AlignDark',DexelaPy.FullWellModes.High,False,False,False)
        el = DexElement.DexElement(False, '',0,1,DexelaPy.bins.x11,37,DexelaPy.ExposureTriggerSource.Internal_Software,0,300,'AlignDark',DexelaPy.FullWellModes.Low,False,False,False)
        RunOneTest(el)

        #el = DexElement.DexElement(False, '',2,1,DexelaPy.bins.x11,162,DexelaPy.ExposureTriggerSource.Internal_Software,45,300,'AlignFlood',DexelaPy.FullWellModes.Low,False,False,False)
        #RunOneTest(el)

        el = DexElement.DexElement(False, '',2,1,DexelaPy.bins.x11,162,DexelaPy.ExposureTriggerSource.Internal_Software,45,300,'AlignFlood',DexelaPy.FullWellModes.High,False,False,False)
        RunOneTest(el)

        if (not easygui.ynbox("Pass Alignment Check?", "Alignment Test", ('Yes', 'No'), None)):
            runSaveData2Server(str(det_model), det_sn_str, rmaNo,False)            
            return False
                            

    ### test sequence file
    cfgFile =getCFGFileName(phase, det_model, bGigE, bSlowMode, buildLabel)
    if cfgFile ==gDefaultCFGFile:
        #warning message has been displayed
        runSaveData2Server(str(det_model), det_sn_str, rmaNo,False)
        return False
                        
    ### clean up old files, if required.
    if os.path.exists("c:\\CMOSTempLog.txt"):
        os.remove("c:\\CMOSTempLog.txt");
                                           
    if os.path.exists("c:\\CMOSTestCSVOutput.csv"):
        os.remove("c:\\CMOSTestCSVOutput.csv");
                        
    if os.path.exists("c:\\CMOSLinearityCSVOutput.csv"):
        os.remove("c:\\CMOSLinearityCSVOutput.csv");
                        
    ### execute test
    DetectorTest(cfgFile, ADCBoardType)

    if bGigE:
        captureSingleDummy()
        
    return True



    
def runPhase2Test(phase, det_model, ADCBoardType, bGigE, bSlowMode, buildLabel):
    easygui.msgbox("Please save the defect map in X-Ray Report Generator before continuing the test!", "Info", 'OK', None, None)

    ### test sequence file
    cfgFile =getCFGFileName(phase, det_model, bGigE, bSlowMode, buildLabel)

    if cfgFile =='none':
        #warning message has been displayed
        return 

    ### execute test
    DetectorTest(cfgFile, ADCBoardType)

    if bGigE:
        captureSingleDummy()

    return



def runSaveData2Server(sModel, sSerialNo, sRMANo, bInitTest):
    if not matchDetectData(dir_output, str(det_model), str(det_sn)):
        pkiWarningBox("Data in Output Directory doesn't match with the detector")
    else:
        ### get numFiles under /Output directory
        onlyfiles = next(os.walk(dir_output))[2] #dir is your directory path as string
        numFiles = len(onlyfiles)

        if numFiles>=1:
            ### copy data to server for both pass/fail detector
            if isRMA(sRMANo):
                rmaData2Server(dir_output, dir_outRMADataServer, str(det_model),  str(det_sn), rmaNo, bInitTest)
            else:
                data2Server(dir_output, dir_outDataServer, str(det_model), det_sn_str)

            ### clean up the data on temp server if any
            dir_src =os.path.join(dir_temp_server, det_sn_str)
            if os.path.exists(dir_src):
                deleteDir(dir_src)

        else:
            pkiWarningBox("No file to transfer")


    return
    


def runSaveData2TempServer(det_sn_str):
    rtnVal =False
    
    ### get numFiles under /Output directory
    onlyfiles = next(os.walk(dir_output))[2] #dir is your directory path as string
    numFiles = len(onlyfiles)

    if numFiles>=1:
        ### save data to temp server.
        dir_dest =os.path.join(dir_temp_server, det_sn_str)
        if os.path.exists(dir_dest):
            deleteDir(dir_dest)

        os.mkdir(dir_dest)
        pkiCopyFiles(dir_output, dir_dest)
            
        rtnVal =True
                
    else:
        pkiWarningBox("No file to transfer")
        rtnVal =False

    return rtnVal



def runDeleteDataOnTempServer(det_sn_str):
    dir_src =os.path.join(dir_temp_server, det_sn_str)
    if os.path.exists(dir_src):
        deleteDir(dir_src)

    return


    

def runGetData(det_model, det_sn_str, ADCBoardType, bGigE, bSlowMode):
    bRMA =False
    rmaNo ='NA'
    bFoundData =False
    
    ### check whehter linearity data is on the server and ready for report
    ###   if failed finding Linearity data, test will stop
    if easygui.ynbox("Is this RMA unit?", "Linearity", ('Yes', 'No'), None):
        bRMA=True
        rmaNo =getRMANo()
    result =getLinearityData(dir_inDataServer, dir_inRMADataServer, dir_UKCofCs, dir_output_parent, str(det_model),det_sn_str, bGigE, bSlowMode, rmaNo, bRMA)
    if not result[0]:
        easygui.msgbox("Linearity Data not found. Test failed and stop here.",
                       "Incoming File Check Error", 'OK', None, None)

        return bFoundData, bRMA, rmaNo

    if easygui.ynbox("Delete old files?", "Directory Clean-up", ('Yes', 'No'), None):
        serverDir =dir_inDataServer
        if bRMA:
            serverDir =dir_inRMADataServer
            
        bFoundData =getLastXRayDataToLocal(serverDir, det_model, det_sn_str, rmaNo, bRMA)
                        
    ### clean up old files, if required.
    if os.path.exists("c:\\CMOSTempLog.txt"):
        os.remove("c:\\CMOSTempLog.txt");
                                           
    if os.path.exists("c:\\CMOSTestCSVOutput.csv"):
        os.remove("c:\\CMOSTestCSVOutput.csv");
                        
    if os.path.exists("c:\\CMOSLinearityCSVOutput.csv"):
        os.remove("c:\\CMOSLinearityCSVOutput.csv");
                        
    return bFoundData, bRMA, rmaNo





def udpateBin22Col128(imagedir):
    return udpateDefectMap_col(imagedir, 128, 16, 1)
    


def udpateDefectMap(theoutdir, filenamepattern ="_Defect_Map_2x2.smv", the_x=128, the_y=-1, old_val=16, new_val=1):
    if os.path.exists(theoutdir):
        print("\n\nupdate defect map under image directory: " +theoutdir)
    else:
        print(theoutdir + " directory not found.")
        return

    ### check the support Microsoft DLL: need be built for the same platform as well, but it is not check here.
    if not os.path.exists("msvcr120d.dll"):
        warn("msvcr120d.dll does not found. Stop Test")
        sys.exit()        
    
    thecmd ="updateDefectMapColLabel.exe /n "

    dirs = os.listdir( theoutdir )
    for thedir in dirs:
        if os.path.isdir(os.path.join(theoutdir,thedir)):
            if "Defect Map" in thedir:
                procdir =os.path.join(theoutdir, "Defect Map")
                maps =os.listdir(procdir)
                for themap in maps:                
                    if filenamepattern in themap:
                        print os.path.join(procdir,themap)
                        os.system(thecmd+"\"" +os.path.join(procdir,themap)+"\"" \
                                  + " /x " +str(the_x) \
                                  + " /y " +str(the_y)\
                                  + " /o " +str(old_val)\
                                  + " /v " +str(new_val))

            if "Customer Delivery" in thedir:
                procdir =os.path.join(theoutdir, "Customer Delivery\\Support Files")
                if os.path.exists(procdir):
                    maps =os.listdir(procdir)
                    for themap in maps:
                        if filenamepattern in themap:
                            print os.path.join(procdir,themap)
                            #os.system(thecmd+"\"" +os.path.join(procdir,themap)+"\"")
                            os.system(thecmd+"\"" +os.path.join(procdir,themap)+"\"" \
                                      + " /x " +str(the_x) \
                                      + " /y " +str(the_y)\
                                      + " /o " +str(old_val)\
                                      + " /v " +str(new_val))


    return


def udpateDefectMap_addPixel(theoutdir, the_x, the_y, filenamepattern ="_Defect_Map_2x2.smv"):
    return udpateDefectMap(theoutdir, filenamepattern, the_x, the_y, 0, 1)
    
def udpateDefectMap_rmPixel(theoutdir, the_x, the_y, filenamepattern ="_Defect_Map_2x2.smv"):
    return udpateDefectMap(theoutdir, filenamepattern, the_x, the_y, 1, 0)

def udpateDefectMap_col(theoutdir, the_x, old_val, new_val, filenamepattern ="_Defect_Map_2x2.smv"):
    return udpateDefectMap(theoutdir, filenamepattern, the_x, -1, old_val, new_val)

def udpateDefectMap_row(theoutdir, the_y, old_val, new_val, filenamepattern ="_Defect_Map_2x2.smv"):
    return udpateDefectMap(theoutdir, filenamepattern, -1, the_y, old_val, new_val)

def runUpdateDefectMap(procdir):
    if not os.path.exists(procdir):
        print(procdir+" not found")
        sys.exit()

    if not os.path.isdir(procdir):
        print(procdir+" is not directory")
        sys.exit()
              
    thelist =os.path.join(procdir, 'updateList.txt')
    if not os.path.exists(thelist):
        print(thelist+" not found")
        sys.exit()

    ### read the update list
    with open(thelist, "r") as readlines:
        for line in readlines:
            line =line.strip()
            line =line.replace(" ","")
            line =line.replace("\t","")            
            
            if line.find("#") <0 and\
               len(line)>2:
                line =re.split('[;|\n]',line)
                if len(line)<5:
                    print("not enough information, stop!")
                    sys.exit()

                dir_img_output =os.path.join(procdir, line[0]+"-"+line[1])
                #print(line +"\t"+dir_img_output)

                if line[2] =='addPix':
                    udpateDefectMap_addPixel(dir_img_output, int(line[3]), int(line[4]))
                elif line[2] =='col' and len(line)>=6:
                    udpateDefectMap_col(dir_img_output, int(line[3]), int(line[4]), int(line[5]))
                elif line[2] =='row' and len(line)>=6:
                    udpateDefectMap_row(dir_img_output, int(line[3]), int(line[4]), int(line[5]))
                else:
                    warn("Not support action " + line[2])
                    sys.exit()
                    
    return


def runCreateCustomerCD(procdir, theDVDDr):
    if not os.path.exists(procdir):
        print(procdir+" not found")
        sys.exit()

    if not os.path.isdir(procdir):
        print(procdir+" is not directory")
        sys.exit()
              
    thelist =os.listdir(procdir)
    for dirname in thelist:
        thedir =os.path.join(procdir, dirname)
        if os.path.isdir(thedir):
            deliverydir =os.path.join(thedir, "Customer Delivery")
            print(thedir)
            if os.path.exists(deliverydir):
                if easygui.ynbox("Get CD ready for "+dirname, "Create CMOS Customer CD", ('Yes', 'No'), None):
                    formatDisk(theDVDDr, 0)
                    data2Dvd(deliverydir, theDVDDr)
        else:
            print(thedir+" is not the directory")
                    
    return
                    
                    

def runDebug():
    print("runDebug...")
    printCofC(dir_output_testRecord)
    sys.exit()

    rmaData2Server(dir_output, dir_outRMADataServer, '1512', '13067', '70124746', True)
    sys.exit()

    #uploadCofCLinearityData("C:\\CMOS\\TestProg\\FaxitronCabinet_toBeReleased\\Linearity", "C:\\CMOS\\TestProg\\FaxitronCabinet_toBeReleased\\Output\\Test Record")
    #sys.exit()
    #sumCMOSDetectorImageIntensity(dir_outDataServer)
    
    #print("return " + str(checkDataOnDvd("D:")))

    #captureSingleDummy()

    #print("%.2f" %runMeasTemperature())
    #DefectCorrectSysNoiseImages(dir_output)

    detMod ="1512"
    detSN = "30461"
    rmaNo ="70122980"

    detMod ="1207"
    detSN = "1224"
    detSN = "01224"
    rmaNo ="BNB02221"
    
    bGigE =False
    bSlowMode =False
    bRMAUnit=True
    dir_scprod_root =dir_inDataServer
    dir_rma_root =dir_inRMADataServer
    dir_uk_root =dir_UKCofCs
    getLinearityData(dir_scprod_root, dir_rma_root, dir_uk_root, dir_output_parent, detMod, detSN, bGigE, bSlowMode, rmaNo, bRMAUnit)
    #getLinearityData(dir_report_root, dir_output_parent, detMod, detSN, bGigE, bSlowMode, rmaNo, bRMAUnit)
    #getLinearityData(dir_report_root, dir_output_parent, detMod, detSN, bGigE, bSlowMode, rmaNo, bRMAUnit)

    sys.exit()
    return



### main
try:
    global cab
    cab = None

    getCurrentDir()

    #runDebug()

    topSelection =None
    while topSelection!='Exit':
        topSelection =easygui.buttonbox(None, 'Main Test',
                                     ('Detector Test', \
                                      'X-Ray System Check', \
                                      'Off System Test',\
                                      'Exit')
                                        )
        print("will run " +topSelection)

        ### init varialbes to default values.
        cfgFile ='none'
        dir_cfg_model =dir_cfg

        det_model =gDefaultVal_detModel
        det_sn =gDefaultVal_detSN    
        det_sn_str =gDefaultVal_detSNStr
        rmaNo =gDefaultVal_RMANo
        bRMA =isRMA(rmaNo)
        buildLabel=gDefaultVal_buildLabel
        buildRev =gDefaultVal_buildRev
        ADCBoardType=gDefaultVal_ADCBoardType
        ADCFirmRev =gDefaultVal_ADCFirmRev
        DAQFirmRev =gDefaultVal_DAQFirmRev
        bGigE=False
        bSlowMode=False
        bMTFHighResolution=False
        bNewGigENoiseSpec =False


        if topSelection == 'Detector Test':

            detTestSelection =None
            while isDefectorConnected() and \
                  (detTestSelection !='Back To Main'):

                ### get detector info
                detTestInfo =getDetTestInfo()
                foundDetector=detTestInfo[gDetInfoIndex_found]
                if foundDetector:
                    det_model =detTestInfo[gDetInfoIndex_model]
                    det_sn =detTestInfo[gDetInfoIndex_serialNo]
                    det_sn_str =detTestInfo[gDetInfoIndex_serialNoStr]
                    rmaNo =detTestInfo[gDetInfoIndex_rmaNo]
                    bRMA =isRMA(rmaNo)
                    buildLabel=detTestInfo[gDetInfoIndex_buildLabel]
                    buildRev =detTestInfo[gDetInfoIndex_buildRev]
                    ADCBoardType=detTestInfo[gDetInfoIndex_ADCBoardType]
                    ADCFirmRev =detTestInfo[gDetInfoIndex_ADCFirmRev]
                    DAQFirmRev =detTestInfo[gDetInfoIndex_DAQFirmRev]
                    bGigE=detTestInfo[gDetInfoIndex_isGigE]
                    bSlowMode=detTestInfo[gDetInfoIndex_isSlowMode]
                    bMTFHighResolution=detTestInfo[gDetInfoIndex_MTFResolution]
                    bNewGigENoiseSpec =detTestInfo[gDetInfoIndex_GigENoisSpec]

                    dir_cfg_model =dir_cfg + str(det_model) +'\\'

                    if  ConfirmDetectorSerialNum(det_sn):

                        detTestSelection =easygui.buttonbox(None, 'Detector Test',\
                                                            ('Production Test',\
                                                             'Init Test',\
                                                             'Eng',\
                                                             'Back To Main',\
                                                             'Exit'))


                        if "30762" in det_sn_str or "30507" in det_sn_str:
                            pkiWarningBox("Unit 30762 and 30507 are engineer units. Please stop the test, hold and notify engineer")
                            detTestSelection='Back To Main'
                            
                        

                        _iflag_quit =-1                 ### quit
                        _iflag_start_test  =0           ### start phase1
                        _iflag_done_phase1 =1           ### done phase1
                        _iflag_done_phase2 =2           ### done phase2
                        _iflag_done_data_transfer =3    ### done data transfer

                        _iflag_start_init_test =11
                        _iflag_done_init_test =12
                        ### the other test don't change value
                        iInTestFlag = _iflag_start_test

                        if detTestSelection =='Init Test' and \
                           isSameDetector(det_model, det_sn): #in case of the detector is removed
                            if easygui.ynbox("Have the previouse data been transferred?", "Data Transfer Reminder:", ('Yes', 'No'), None):
                                iInTestFlag = _iflag_start_init_test

                                if runPhase1Test("Init Test", det_model, det_sn_str, rmaNo, ADCBoardType, bGigE, bSlowMode, buildLabel): #if 'False', data has been transferred.
                                    runSaveData2Server(str(det_model), det_sn_str, rmaNo, True)
                                
                                iInTestFlag = _iflag_done_init_test

                            ### back to main menu
                            detTestSelection ='Back To Main'
                            
                        elif detTestSelection =='Production Test' and \
                             isSameDetector(det_model, det_sn): #in case of the detector is removed

                            linearDataServer =dir_inDataServer
                            bFoundLinearityData =False

                            productionTestSelection =None

                            ### get linearity data
                            result =getLinearityData(dir_inDataServer, dir_inRMADataServer, dir_UKCofCs, dir_output_parent, str(det_model),det_sn_str, bGigE, bSlowMode, rmaNo, bRMA)
                            
                            bFoundLinearityData=result[0]

                            if bFoundLinearityData:
                                bRunDetectorTest =True
                            else: #allow to run test without linearity data.
                                bRunDetectorTest =(easygui.ynbox("Linearity Data not found, and cannot create CofC. Continue?",\
                                                  "Linearity Data Check",\
                                                  ('Yes', 'No'),\
                                                  None))
                                

                            while isSameDetector(det_model, det_sn) and \
                                  bRunDetectorTest and \
                                  productionTestSelection !='Back To Main':

                                productionTestSelection =easygui.buttonbox(None, 'Detector Test',
                                                                           ('Phase 1', \
                                                                            'Phase 2', \
                                                                            'Test Summary',\
                                                                            'Re-take Phantom Images', \
                                                                            'Save Data', \
                                                                            'Get Data',\
                                                                            'Load Image for Review',\
                                                                            'Alignment', \
                                                                            'Snap One',\
                                                                            'Eng', \
                                                                            'Back To Main'))


                                if productionTestSelection =='Phase 1':
                                    if easygui.ynbox("Have the previouse data been transferred?", "Data Transfer Reminder:", ('Yes', 'No'), None):
                                        iInTestFlag = _iflag_start_test

                                        bContinueTest =True
                                        
                                        if buildLabel =='2315N2-C22-HECC':
                                            measedT =float(runMeasTemperature())

                                            if measedT < 20.0 and measedT >45.0:
                                                bContinueTest=False
                                                pkiWarningBox("Failed temperature test: meased %.2f spec: 20.0~45.0C" %measedT)

                                        if bContinueTest:
                                            runPhase1Test(productionTestSelection, det_model, det_sn_str, rmaNo, ADCBoardType, bGigE, bSlowMode, buildLabel)

                                        iInTestFlag = _iflag_done_phase1

                                    
                                elif productionTestSelection =='Phase 2':
                                    ###specialHandlingSysNoiseFailure- DefectCorrectSysNoiseImages(dir_output)
                                    if iInTestFlag == _iflag_done_phase1 or \
                                       easygui.ynbox("The unit must finish 'Phase 1' test before this test. \nDo you want to continue this test?", \
                                                     "Phase2 confirmation", ('Yes', 'No'), None):

                                        runPhase2Test(productionTestSelection, det_model, ADCBoardType, bGigE, bSlowMode, buildLabel)

                                        iInTestFlag = _iflag_done_phase2


                                    
                                elif productionTestSelection =='Test Summary':
                                    
                                    if not matchDetectData(dir_output, str(det_model), str(det_sn)):
                                        pkiWarningBox("Data in Output Directory doesn't match with the detector")
                                    else: #matchDetectorData
                                        renameTestRecords(dir_output_testRecord, buildLabel, det_sn_str)
                                        
                                        if easygui.ynbox("Test Result", "Test Summary", ('Pass', 'Fail'), None): #'Pass' return True; 'Fail' return False
                                            if not bFoundLinearityData:
                                                pkiWarningBox("Linearity Data not Found. This is only transfer data for pass unit. Use 'Failed Test' or 'Save Data' instead.")
                                            else:
                                                if iInTestFlag ==_iflag_done_phase2 or \
                                                   iInTestFlag ==_iflag_done_data_transfer or \
                                                   easygui.ynbox("Unit should finish phase 1&2 test before transfer data.\nIs this a special transfer? If yes, try again. It will perform data transfer.", \
                                                                   "Special Data Transfer", ('Yes', 'No'), None): ### allow to transfer more than once
                                                    ### update CofC with linearity data, and get test result from data record.
                                                    ###   if data record pass, --> get visual inspection result from the user

                                                    ### Correct the images with the latest defectMap
                                                    DefectCorrect(dir_output)
                                                    #noNeed.TheCustomerDeliveryDirWillBeReCreated- DefectCorrect(dir_output_custDelivery_TestImages)

                                                    if "1512-C-3D-C600B" in buildLabel:
                                                        udpateBin22Col128(dir_output)


                                                    bRelease =False
                                                    result =updateCofCLinearityData(os.path.join(dir_output_parent, 'Linearity'),
                                                                                    dir_output_testRecord, #no need serial number. It is recorded in CofC already.
                                                                                    temp_csv_output,
                                                                                    buildLabel,
                                                                                    buildRev,
                                                                                    ADCFirmRev,
                                                                                    DAQFirmRev,
                                                                                    bSlowMode,
                                                                                    bMTFHighResolution,
                                                                                    bNewGigENoiseSpec,
                                                                                    bRMA)

                                                    if result[0]:
                                                        easygui.msgbox("Data Record Show: "+result[1], "Test Result from Test Data", 'OK', None, None)

                                                        if result[1]=="PASS":
                                                            bRelease =True
                                                    else:
                                                        bRelease =False
                                                        pkiWarningBox("Failed update CofC. Test Failed")


                                                    ### copy phantom images to customer delivery directory
                                                    if bRelease:
                                                        ### recreate customer delivery folders(Suppor Filese/Image Files folder)
                                                        print("\n\n\ngoing to copyImages2Delivery")
                                                        if not copyImages2Delivery(dir_output, dir_output_AutoCal, dir_output_Darks, dir_output_Floods, dir_output_DefectMap,\
                                                                                   dir_output_custDelivery, dir_output_custDelivery_TestImages, dir_output_custDelivery_SupportFiles):
                                                            pkiWarningBox("Error found with " +dir_output_custDelivery+ ". Please check before create CD")

                                                        ### Display Final test result, and create CD for passed detector
                                                        easygui.msgbox("PASS", "Final Test Result", 'OK', None, None)

                                                        if gbReleaseCD:
                                                            if easygui.ynbox("Create CD", "Create CD?", ('Yes', 'No'), None):
                                                                easygui.msgbox("Insert a blank CD to attached CD drive and wait for the next instruction", "Alert", 'OK', None, None)
                                                                time.sleep(2) #wait for 2s to allow CD
                                                                #useFollowing- pkiWaitMessage("Waiting for CD getting ready...", 0.1, 10)
                                                                count=0
                                                                while (not easygui.ynbox("Press Yes when disc is ready.", "Alert") and count<10):
                                                                    print("wait.... Press 'Yes' to continue when it is ready");
                                                                    time.sleep(5)
                                                                    count +=1

                                                                print("CD is ready and proceed to copy data to CD")
                                                                if not isDiskReady(drv_dvd, str(det_model)+'_'+det_sn_str):
                                                                    pkiWarningBox("CD is not ready. Please check")
                                                                    #if not data2Dvd(dir_output_custDelivery, drv_dvd):
                                                                    #    pkiWarningBox("Failed copy data to CD")
                                                                else:
                                                                    data2Dvd(dir_output_custDelivery, drv_dvd)
                                                    else:
                                                        easygui.msgbox("FAIL", "Final Test Result", 'OK', None, None)
                                                                            

                                                    ### copy data to server for both pass/fail detector
                                                    if bRMA:
                                                        rmaData2Server(dir_output, dir_outRMADataServer, str(det_model),  str(det_sn), rmaNo, False)
                                                    else:
                                                        data2Server(dir_output, dir_outDataServer, str(det_model), det_sn_str)

                                                    if bRelease:
                                                        printCofC(dir_output_testRecord)
                                                    else:
                                                        pkiWarningBox("Data transfer failed. No CofC printed.\nPlease save the Data mannually. And call Eng support")
                                                        

                                                    ### load report to DB
                                                    if gbLoadDB and os.path.exists(dir_output+"\\Test Report"):
                                                        det_mod_str=str(det_model)
                                                        if rmaNo==None:
                                                            print det_model, det_sn_str, 'NA', buildLabel, dir_output
                                                            loadReportToDB(det_mod_str, det_sn_str, 'NA', buildLabel, dir_output)
                                                        else:
                                                            print det_model, det_sn_str, rmaNo, buildLabel, dir_output
                                                            loadReportToDB(det_mod_str, det_sn_str, rmaNo, buildLabel, dir_output)

                                                    ### clean up the data on temp server if any
                                                    runDeleteDataOnTempServer(det_sn_str)

                                                    easygui.msgbox("File Transfer Completed", "File Transfer Done", 'OK', None, None)

                                                    iInTestFlag = _iflag_done_data_transfer
                                                    
                                            productionTestSelection ='Back To Main'
                                            detTestSelection ='Back To Main'
                                                    
                                                #elif easygui.ynbox("Unit should finish phase 1&2 test before transfer data.\nIs this a special transfer? If yes, try again. It will perform data transfer.", \
                                                #                   "Special Data Transfer", ('Yes', 'No'), None):
                                                #    iInTestFlag =_iflag_done_phase2 #set to done "phase 2" flag, to enable data transfer.

                                        else: #Failed Test
                                            ### get numFiles under /Output directory
                                            onlyfiles = next(os.walk(dir_output))[2] #dir is your directory path as string
                                            numFiles = len(onlyfiles)

                                            if numFiles>=1:
                                                ### copy data to server for both pass/fail detector
                                                #if not data2Server(dir_output, dir_outDataServer, str(det_model), det_sn_str):
                                                #    pkiWarningBox("Failed copy data to server")
                                                if bRMA:
                                                    rmaData2Server(dir_output, dir_outRMADataServer, str(det_model),  str(det_sn), rmaNo, False)
                                                else:
                                                    data2Server(dir_output, dir_outDataServer, str(det_model), det_sn_str)

                                                ### clean up the data on temp server if any
                                                dir_src =os.path.join(dir_temp_server, det_sn_str)
                                                if os.path.exists(dir_src):
                                                    deleteDir(dir_src)


                                                ### load report to DB
                                                if gbLoadDB and os.path.exists(dir_output+"\\Test Report"):
                                                    det_mod_str=str(det_model)
                                                    if rmaNo==None:
                                                        print det_model, det_sn_str, 'NA', buildLabel, dir_output
                                                        loadReportToDB(det_mod_str, det_sn_str, 'NA', buildLabel, dir_output)
                                                    else:
                                                        print det_model, det_sn_str, rmaNo, buildLabel, dir_output
                                                        loadReportToDB(det_mod_str, det_sn_str, rmaNo, buildLabel, dir_output)

                                                easygui.msgbox("File Transfer Completed", "File Transfer Done", 'OK', None, None)

                                                ### clean up the data on temp server if any
                                                runDeleteDataOnTempServer(det_sn_str)
                                                
                                                iInTestFlag = _iflag_done_data_transfer

                                            else:
                                                pkiWarningBox("No file to transfer")

                                            productionTestSelection ='Back To Main'
                                            detTestSelection ='Back To Main'


                                elif productionTestSelection =='Re-take Phantom Images':
                                    #if (iInTestFlag ==_iflag_done_phase1 or iInTestFlag ==_iflag_done_phase2):
                                    if True:

                                        imageSel =easygui.buttonbox('Select Image.', 'Retake Phantom Images',
                                                                    ('Phase 1 hand-phantom', \
                                                                     'Phase 2 hand-phantom', \
                                                                     'digimam-phantom', \
                                                                     'CIRS-mesh-phantom', \
                                                                     'TOR-MAS-phantom', \
                                                                     'MTF-phantom',\
                                                                     'AlBar-phantom'))

                                        ### test sequence file
                                        cfgFile =getCFGFileName("Phase 2", str(det_model), bGigE, bSlowMode, buildLabel)                            

                                        phantomName ="none"
                                        if imageSel =="Phase 1 hand-phantom":
                                            cfgFile =getCFGFileName("Phase 1", det_model, bGigE, bSlowMode, buildLabel)
                                            phantomName ="hand-phantom"
                                        elif imageSel =="Phase 2 hand-phantom":
                                            #cfgFile =getCFGFileName('Phase 2', det_model, bGigE, bSlowMode, buildLabel)
                                            phantomName ="hand-phantom"
                                        elif imageSel =="digimam-phantom":
                                            #cfgFile =getCFGFileName('Phase 2', det_model, bGigE, bSlowMode, buildLabel)
                                            phantomName ="digimam-phantom"
                                        elif imageSel =="CIRS-mesh-phantom":
                                            #cfgFile =getCFGFileName('Phase 2', det_model, bGigE, bSlowMode, buildLabel)
                                            phantomName ="CIRS-mesh-phantom"
                                        elif imageSel =="TOR-MAS-phantom":
                                            #cfgFile =getCFGFileName('Phase 2', det_model, bGigE, bSlowMode, buildLabel)
                                            phantomName ="TOR-MAS-phantom"
                                        elif imageSel =="MTF-phantom":
                                            #cfgFile =getCFGFileName('Phase 2', det_model, bGigE, bSlowMode, buildLabel)
                                            phantomName ="MTF-phantom"
                                        elif imageSel =="AlBar-phantom":
                                            #cfgFile =getCFGFileName('Phase 2', det_model, bGigE, bSlowMode, buildLabel)
                                            phantomName ="AlBar-phantom"
                                        else:
                                            easygui.msgbox("Not support phantom image. Test stopped!", "Warning...", 'OK', None, None)


                                        ### execute test
                                        #print("run test " + cfgFile + ", " + ADCBoaryType+ ", " + phantomName)
                                        DetectorTest(cfgFile, ADCBoardType, phantomName)
                                        
                                    else:
                                        pkiWarningBox("The unit must finish 'Phase 1' test before this test")  




                                elif productionTestSelection =='Save Data':
                                    if runSaveData2TempServer(det_sn_str):
                                        iInTestFlag = _iflag_done_data_transfer


                                elif productionTestSelection =='Get Data':
                                    ### get the temp data from server to local for additional test.
                                    ### this is created for phase1 and phase2 will be done for 2-connection at the 2 different time.
                                    ###     After phase1, the data should be saved on the temp server.
                                    bFoundData =False

                                    if easygui.ynbox("Have the previouse data been transferred?", "Data Transfer Reminder:", ('Yes', 'No'), None):
                                        iInTestFlag = _iflag_start_test

                                        ### search data on the temp server
    
                                        if bFoundData:
                                            ### clean up local dir
                                            if easygui.ynbox("Delete old files?", "Directory Clean-up", ('Yes', 'No'), None):
                                                prepDir(dir_output)

                                            ### copy the phase1 data to local
                                            dir_src =os.path.join(dir_temp_server, det_sn_str)
                                            if os.path.exists(dir_src):
                                                print("src dir:"+dir_src)
                                                pkiCopyFiles(dir_src, dir_output)
                                            else:
                                                bFoundData =False
                                                pkiWarningBox("No Data not found. Test failed and stop here.")


                                        iInTestFlag = _iflag_done_phase1
                                    else:
                                        pkiWarningBox("Save the old file. Then run this again")


                                elif productionTestSelection =='Load Image for Review':
                                    startReportApp(dir_output)

                                    
                                elif productionTestSelection =='Alignment':
                                    cnt =0

                                    '''
                                    while (15+cnt*5)<100:
                                        print "setting... DexElement."
                                        #el = DexElement.DexElement(isCommand, command,xOnTimeSecs,numberOfExposures,Binning,t_expms,trigger,kVp,mA,comment,fullWell,darkCorrect,gainCorrect,defectCorrect)
                                        el = DexElement.DexElement(False, '',0,1,DexelaPy.bins.x11,70,DexelaPy.ExposureTriggerSource.Internal_Software,0,300,'AlignDark'+str(cnt),DexelaPy.FullWellModes.High,False,False,False)
                                        RunOneTest(el, False, False)
                                        el = DexElement.DexElement(False, '',0,1,DexelaPy.bins.x11,10038,DexelaPy.ExposureTriggerSource.Internal_Software,0,300,'AlignDark'+str(cnt),DexelaPy.FullWellModes.High,False,False,False)
                                        RunOneTest(el, False, False)
                                        el = DexElement.DexElement(False, '',0,1,DexelaPy.bins.x11,10038,DexelaPy.ExposureTriggerSource.Internal_Software,0,300,'AlignDark'+str(cnt),DexelaPy.FullWellModes.High,True,True,False)
                                        RunOneTest(el, False, False)

                                        el = DexElement.DexElement(False, '',2,1,DexelaPy.bins.x11,162,DexelaPy.ExposureTriggerSource.Internal_Software,15+cnt*5,300,'AlignFlood'+str(cnt),DexelaPy.FullWellModes.Low,False,False,False)
                                        RunOneTest(el, False, False)
                                        time.sleep(10)
                                        el = DexElement.DexElement(False, '',2,1,DexelaPy.bins.x11,162,DexelaPy.ExposureTriggerSource.Internal_Software,15+cnt*5,300,'AlignFlood'+str(cnt),DexelaPy.FullWellModes.Low,True,True,False)
                                        RunOneTest(el, False, False)
                                        

                                        time.sleep(20)
                                        cnt =cnt+1
                                    '''


                                    while cnt<60:
                                        print "setting... DexElement."
                                        #el = DexElement.DexElement(isCommand, command,xOnTimeSecs,numberOfExposures,Binning,t_expms,trigger,kVp,mA,comment,fullWell,darkCorrect,gainCorrect,defectCorrect)
                                        el = DexElement.DexElement(False, '',0,1,DexelaPy.bins.x11,37,DexelaPy.ExposureTriggerSource.Internal_Software,0,300,'AlignDark'+str(cnt),DexelaPy.FullWellModes.High,False,False,False)
                                        RunOneTest(el, False, False)
                                        el = DexElement.DexElement(False, '',0,1,DexelaPy.bins.x11,10038,DexelaPy.ExposureTriggerSource.Internal_Software,0,300,'AlignDark'+str(cnt),DexelaPy.FullWellModes.High,False,False,False)
                                        RunOneTest(el, False, False)

                                        el = DexElement.DexElement(False, '',2,1,DexelaPy.bins.x11,162,DexelaPy.ExposureTriggerSource.Internal_Software,45,300,'AlignFlood'+str(cnt),DexelaPy.FullWellModes.Low,False,False,False)
                                        RunOneTest(el, False, False)
                                        

                                        time.sleep(5*60)
                                        cnt =cnt+1

                                    runSaveData2Server(str(det_model), det_sn_str, rmaNo, False)
                                    
							
                                elif productionTestSelection =='Back To Main':
                                    detTestSelection = 'Back To Main'

                                    
                                else:
                                    pkiWarningBox("Note support production test selection: " + productionTestSelection)
                                    
                                

                        elif detTestSelection =='Eng':

                            engTestSelection =None

                            while (engTestSelection !='Back To Main') and \
                                  isSameDetector(det_model, det_sn): #in case of the detector is removed
                                
                                engTestSelection =easygui.buttonbox(None, 'Detector Test',
                                                                    ('Snap One',\
                                                                     'Eng Flow',\
                                                                     'Meas. Temperature',\
                                                                     'Back To Main'))

                                if engTestSelection =='Meas. Temperature':
                                    pkiInfoBox(str("Detector Temperature %.2f" %RunReadTempInC()))


                                elif engTestSelection =='Eng Flow':
                                    engCfgFile =str(easygui.fileopenbox("Image Directory","Image Directory", None))
                                    pkiInfoBox(engCfgFile)

                                    DetectorTest(engCfgFile, ADCBoardType)
                                    
                                elif engTestSelection =='Snap One':
                                    theUseCommand =False
                                    theCommand  =''
                                    theDarkCorrect =False
                                    theGainCorrect =False
                                    theDefCorrect =False
                                    theBinning =DexelaPy.bins.x11
                                    theWell=DexelaPy.FullWellModes.High
                                    theNumFrames =1

                                    theSettings =getOneImageAcqSetting()
                                    index =0; print("index =%d" % index)
                                    theKVp =int(theSettings[index])

                                    index +=1; print("index =%d" % index)
                                    theXRayOnTime =int(theSettings[index])

                                    index +=2; print("index =%d" % index)
                                    if theSettings[index]=="Y":
                                        theUseCommand =True
                                    else:
                                        theUseCommand =False

                                    index +=1; print("index =%d" % index)
                                    if theUseCommand:
                                        theCommand =theSettings[index]
                                    else:
                                        theCommand =''

                                    index +=1; print("index =%d" % index)
                                    theNumFrames =int(theSettings[index])

                                    index +=1; print("index =%d" % index)
                                    if theSettings[index]==1:
                                        theBinning =DexelaPy.bins.x11
                                    elif theSettings[index]==2:
                                        theBinning =DexelaPy.bins.x22
                                    else: #theSettings[index]==4:
                                        theBinning =DexelaPy.bins.x44

                                    index +=1;print("index =%d" % index)
                                    theExpTime =int(theSettings[index])

                                    index +=1;print("index =%d" % index)
                                    if theSettings[index]=="H":
                                        theWell =DexelaPy.FullWellModes.High
                                    else:
                                        theWell =DexelaPy.FullWellModes.Low

                                    index +=1;print("index =%d" % index)
                                    if theSettings[index]=="Y":
                                        theDarkCorrect =True
                                    else:
                                        theDarkCorrect =False

                                    index +=1;print("index =%d" % index)
                                    if theSettings[index]=="Y":
                                        theGainCorrect =True
                                    else:
                                        theGainCorrect =False

                                    index +=1;print("index =%d" % index)
                                    if theSettings[index]=="Y":
                                        theDefCorrect =True
                                    else:
                                        theDefCorrect =False

                                    print("Snap One")
                                    el =DexElement.DexElement(theUseCommand, theCommand,\
                                                              theXRayOnTime,
                                                              theNumFrames,theBinning,theExpTime,DexelaPy.ExposureTriggerSource.Internal_Software,\
                                                              theKVp,300,\
                                                              'AlignFlood',\
                                                              theWell,\
                                                              theDarkCorrect,theGainCorrect,theDefCorrect)
                                    RunOneTest(el, False, False)

                                
                        elif detTestSelection == 'Back To Main':
                            print(detTestSelection)

                        elif detTestSelection =='Exit':
                            print(detTestSelection)
                            sys.exit()

                        
                
        elif topSelection =='X-Ray System Check':
            sysTestSelection =None
            while sysTestSelection!='Back To Main':
                sysTestSelection =easygui.buttonbox(None, 'X-Ray system Check',
                                                    ('Fire X-Ray', \
                                                     'X-Ray Cab. Dosi. Meas.',\
                                                     'Back To Main'))
                print("will run " + sysTestSelection)

                if sysTestSelection =='Fire X-Ray':
                    xraysetting =getXRaySetting()
                    kVpSetting =int(xraysetting[0])
                    XRayOnTime =int(xraysetting[1])
                    
                    bFireXRay =True

                    while bFireXRay:
                        cab = Faxitron.Cabinet(3)

                        print "Fire X-Ray for ", XRayOnTime, "sec at ", kVpSetting, " kVp"

                        if(kVpSetting != 0):
                            cab.Configure(kVpSetting,XRayOnTime*1000)
                            if cab.FireXRay() == True:
                                cab.WaitXRayOn()

                            if(kVpSetting != 0):
                                time.sleep(5) #need have some wait time here. Otherwise it will cause exception.
                            cab.WaitXRayOff()

                            if cab != None:
                                cab.close()

                            print "X-Ray Off"
                        bFireXRay =easygui.ynbox("Fire X-Ray", "Continue Fire X-Ray?", ('Yes', 'No'), None)


                elif sysTestSelection =='X-Ray Cab. Dosi. Meas.':
                    xraysetting =getXRaySetting()
                    kVpSetting =int(xraysetting[0])
                    XRayOnTime =int(xraysetting[1])

                    bInCabMeas =True
                    readingCtr =None
                    readingBL =None
                    readingTL =None
                    readingTR =None
                    readingBR =None
                    
                    while bInCabMeas:
                        meassopt =easygui.buttonbox('Select Meas. Spot.\nMeased:\n'+\
                                                    '\tCenter ='+str(readingCtr)+'\n'\
                                                    '\treadingBL\t ='+str(readingBL)+'\n'\
                                                    '\treadingTL\t ='+str(readingTL)+'\n'\
                                                    '\treadingTR\t ='+str(readingTR)+'\n'\
                                                    '\treadingBR\t ='+str(readingBR)+'\n'\
                                                    , 'X-Ray Cab Dosi. Meas.',
                                                    ('Center',
                                                     'Bottom-Left',
                                                     'Top-Left',
                                                     'Top-Right',
                                                     'Bottom-Right',
                                                     'Done Meas.'))

                        if meassopt == 'Done Meas.':
                            bInCabMeas=False

                            if readingCtr ==None or \
                               readingBL ==None or \
                               readingTL ==None or \
                               readingTR == None or\
                               readingBR == None:
                                bInCabMeas =pkiYNBox("Not cover all 5-spot yet. Continue measurement?", "X-Ray Cab Dosi Meas.")

                        else:
                            goFireXRay(kVpSetting, XRayOnTime)
                            dosireading =easygui.enterbox('Enter Dosimeter reading in mR (number only)', 'Dosimeter Reading', '', True)

                            if meassopt =='Center':
                                print(str(dosireading))                
                                readingCtr =str(dosireading)
                            elif meassopt =='Bottom-Left':
                                readingBL =str(dosireading)
                            elif meassopt =='Top-Left':
                                readingTL =str(dosireading)
                            elif meassopt =='Top-Right':
                                readingTR =str(dosireading)
                            elif meassopt =='Bottom-Right':
                                readingBR =str(dosireading)


                    ### record to file
                    f =None
                    if (os.path.isfile(XRayCabDosiLog)):
                        f =open(XRayCabDosiLog, 'a')
                    else:
                        f =open(XRayCabDosiLog, 'w+')
                        f.write('#time, center, Bottom-Left, Top-Left, Top-Right, Bottom-Right (in mR)\n')
                        
                    f.write('%s.,%s,%s,%s,%s,%s\n' % ( datetime.date.today(),str(readingCtr),readingBL,readingTL,readingTR,readingBR))
                        
                    if f!=None:
                        f.close()
                            
                        
                    

        elif topSelection=='Off System Test':
            offlineTestSelection =None
            while offlineTestSelection!='Back To Main':
                offlineTestSelection =easygui.buttonbox(None, 'Image Check',
                                                    ('CheckData', \
                                                     'Eng Data Transfer',\
                                                     'Image Defect Correction', \
                                                     'Hologic Correction',\
                                                     'Update CofC',\
                                                     'CD from Server',\
                                                     'Load Offline Image',\
                                                     'Back To Main'))
                print("will run " + offlineTestSelection)

                if offlineTestSelection=='CheckData':
                    printImgBlockAvg()


                elif offlineTestSelection=='Eng Data Transfer':
                    data2EngServer(dir_output, dir_dataEngServer, 'Output', True, str(det_model), det_sn_str)
                    data2EngServer(dir_linear, dir_dataEngServer, 'Linearity', False, str(det_model), det_sn_str)


                elif offlineTestSelection=='Image Defect Correction':
                    imagedir =str(easygui.diropenbox("Image Directory","Image Directory", None))
                    DefectCorrect(imagedir)
                    #DefectCorrect(dir_output)


                elif offlineTestSelection=='Hologic Correction':
                    ### Binx22 defect map correction:
                    procdir =str(easygui.diropenbox("Directory","Hologic Correction", None))
                    runUpdateDefectMap(procdir)


                    if pkiYNBox("Create CDs", "New CDs"):
                        theDVDDr =easygui.choicebox('Select DVD/CD drive', 'DVD/CD drive List', ('D:', 'E:','F:'))#,('OK', 'Cancel'))
                        if theDVDDr ==None:
                            pkiWarningBox('No DVD/CD drive selected. Skip create CD');
                        else:
                            runCreateCustomerCD(procdir, theDVDDr)

                elif offlineTestSelection =='Update CofC':
                    thedetinfo =getOfflineDetInfo()
                    
                    det_model =thedetinfo[gDetInfoIndex_model]
                    det_sn =thedetinfo[gDetInfoIndex_serialNo]
                    det_sn_str =thedetinfo[gDetInfoIndex_serialNoStr]
                    rmaNo =thedetinfo[gDetInfoIndex_rmaNo]
                    buildLabel=thedetinfo[gDetInfoIndex_buildLabel]
                    buildRev =thedetinfo[gDetInfoIndex_buildRev]
                    ADCBoardType=thedetinfo[gDetInfoIndex_ADCBoardType]
                    ADCFirmRev =thedetinfo[gDetInfoIndex_ADCFirmRev]
                    DAQFirmRev =thedetinfo[gDetInfoIndex_DAQFirmRev]
                    bGigE=thedetinfo[gDetInfoIndex_isGigE]
                    bSlowMode=thedetinfo[gDetInfoIndex_isSlowMode]
                    bMTFHighResolution=thedetinfo[gDetInfoIndex_MTFResolution]
                    bNewGigENoiseSpec =thedetinfo[gDetInfoIndex_GigENoisSpec]

                    print rmaNo
                    print gDefaultVal_RMANo
                    if rmaNo ==gDefaultVal_RMANo:
                        bRMA =False
                        print("not RMA")
                    else:
                        bRMA =True
                        print("RMA")

                    ### check if images/test recrods under 'Output' directory is for the to-be-updadted units
                    ###     'Yes' - move to the next step
                    ###     'No' - copy the data to 'Output' directory

                    ### check if need copy the linearity data to local

                    '''
                    msg = "Enter info for the detector"
                    title = "Detector Info"
                    fieldNames = ["\t Serial No.:","\t RMA No.:"]
                    fieldValues = ['', '']  # we start with defaults for the values
                    fieldValues = easygui.multenterbox(msg,title, fieldNames,fieldValues)
                    while not (pkiYNBox("Is the correct info?\n\tSerial No:. " +str(fieldValues[0]) +"\n\tRMA No.:" + str(fieldValues[1]), "Detector Info confirmation")):
                        fieldValues = easygui.multenterbox(msg,title, fieldNames,fieldValues)

                    '''

                    '''
                    updateCofCLinearityData(os.path.join(dir_output_parent, 'Linearity'),
                                                    dir_output_testRecord, #no need serial number. It is recorded in CofC already.
                                                    temp_csv_output,
                                                    '1512N-G16-DRZS',
                                                    'B',
                                                    '31C',
                                                    '3',
                                                    'N',
                                                    'L',
                                                    'Y',
                                                    True)                       
                    sys.exit()
                    '''

                    
                    result =updateCofCLinearityData(os.path.join(dir_output_parent, 'Linearity'),
                                                    dir_output_testRecord, #no need serial number. It is recorded in CofC already.
                                                    temp_csv_output,
                                                    buildLabel,
                                                    buildRev,
                                                    ADCFirmRev,
                                                    DAQFirmRev,
                                                    bSlowMode,
                                                    bMTFHighResolution,
                                                    bNewGigENoiseSpec,
                                                    bRMA)                                      
                    

                elif offlineTestSelection =='CD from Server':
                    offlineDetInfo =getOfflineDet_model_serialNo()
                    sModel =str(offlineDetInfo[1])
                    sSerNo =str(offlineDetInfo[3])
                    sRMANo =str(getRMANo())
                    
                    if offlineDetInfo[0]:
                        easygui.msgbox("Insert a blank CD to attached CD drive and wait for the next instruction", "Alert", 'OK', None, None)
                        time.sleep(2) #wait for 2s to allow CD
                        count=0
                        while (not easygui.ynbox("Press Yes when disc is ready.", "Alert") and count<10):
                            print("wait.... Press 'Yes' to continue when it is ready");
                            time.sleep(5)
                            count +=1


                        theDVDDr =easygui.choicebox('Select DVD/CD drive', 'DVD/CD drive List', ('D:', 'E:','F:'))#,('OK', 'Cancel'))
                        if theDVDDr ==None:
                            pkiWarningBox('No DVD/CD drive selected. Skip create CD');
                        else:
                            if not isDiskReady(theDVDDr, sModel+'_'+sSerNo):
                                pkiWarningBox("CD is not ready. Please check")
                            else:                                
                                rootdir =""
                                filepattern ='*'
                                if sRMANo is None or\
                                   sRMANo =='NA' or \
                                   len(sRMANo)==0:
                                    rootdir =os.path.join(dir_outDataServer,sModel+"-"+sSerNo+"\\X-Ray")
                                else:
                                    rootdir =os.path.join(dir_outRMADataServer, sModel+"\\"+sRMANo+"_"+sSerNo)
                                    filepattern ="*X-Ray*"

                                srcdir =os.path.join(getLatestDir(rootdir, filepattern), "Customer Delivery")
                                if os.path.exists(srcdir):
                                    data2Dvd(srcdir, theDVDDr)
                                else:
                                    pkiWarningBox(rootdir + " doesn't exist. CD is not created. Please check")
                                
                elif offlineTestSelection =='Load Offline Image':
                    offlineDetInfo =getOfflineDet_model_serialNo()
                    sModel =str(offlineDetInfo[1])
                    sSerNo =str(offlineDetInfo[3])
                    sRMANo =str(getRMANo())
                    
                    if offlineDetInfo[0]:
                        ### get the temp data from server to local for additional test.                     

                        if easygui.ynbox("Have the previouse data been transferred? Delete Output files?", "Data Transfer Reminder:", ('Yes', 'No'), None):
                            #iInTestFlag = _iflag_start_test
                            prepDir(dir_output)

                            ### search data on the temp server
                            if getOfflineImages(dir_inDataServer, dir_inRMADataServer, dir_output, sModel,sSerNo, sRMANo, bRMA):
                                startReportApp(dir_output)
                            else:
                                pkiWarningBox("No Data not found for" + str(det_model) + "-" + sModel +"(RMA No.:" + sRMANo+")")

                        else:
                            pkiWarningBox("Save the old file. Then run this again")
                    
            
        else:
            print("No test defined for " +topSelection)
    

    print("All Done. Program exit.")

     
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




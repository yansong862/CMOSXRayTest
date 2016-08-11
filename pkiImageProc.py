from ctypes import *
import time
import copy
import os
import sys
import re
import ctypes
import glob
import filecmp
import ntpath
import datetime

from FileUtils import *
import DexelaPy 


def sumCMOSDetectorImageIntensity(rootDataDir):
    curdir = os.getcwd()
    thescript =os.path.join(curdir, 'bin', 'cmosXray_ListOfBuilds.pl')
    if not os.path.exists(thescript):
        print( thescript + " doesn't exist")

    thecmd ="perl " + thescript

    #os.system(thecmd)

    
    
    
    detListFile ="C:\\CMOSListOfBuild.txt"

    f = open("c:\\CMOSDetectorImageIntensitySummary.csv",'w+')
    f.write("#BuildLabel,Model,Serial,Binning,FW,XRaykVP,XRay_uA,ExpTime,GainCorrect,DefectCorrect,ImageAvg\n")
    
    with open(detListFile, "r") as detlist:
        for line in detlist:
            line.strip()
            if line.find ("#")<0 and len(line)>0: #not comment line
                words =re.split(",", line)
                
                if len(words)>=3:
                    theModel =words[0].strip()
                    theSerNo =words[1].strip()
                    theBuildLabel =words[2].strip()

                    detDir =os.path.join(rootDataDir, theModel+"-"+theSerNo)
                    lastXRayDir =getLatestDir(os.path.join(detDir,"X-Ray"))
                    print lastXRayDir

                    for filename in glob.glob(os.path.join(lastXRayDir, "*.tif")):
                        if filename.find("dark")>=0 or filename.find("flood")>=0:
                            basefilename =os.path.basename(filename)
                            basefilename =basefilename.replace(" ","")
                            print basefilename
                            settings=re.split('[-_]',basefilename)
                            if len(settings)<9: next

                            image =DexelaPy.DexImagePy(filename)

                            
                            theModel =settings[0]
                            theSerNo =settings[1]
                            FW =settings[2]
                            expTime =settings[3]
                            binning =settings[4]
                            GainCorrect =False
                            DefectCorrect=False
                            XRay_kVp ="0"
                            XRay_uA="0"
                            imgavg =str("%.0f" % image.PlaneAvg(0))

                            if image.IsFloodCorrected() and image.IsDefectCorrected():
                                GainCorrect =True
                                DefectCorrect=True
                                XRay_kVp =settings[7]
                                XRay_uA=settings[8]

                                if XRay_kVp.find("kVp")<0:
                                    XRay_kVp =settings[8]
                                    XRay_uA=settings[9]
                                    
                            elif image.IsFloodCorrected():
                                GainCorrect =True
                                DefectCorrect=False
                                XRay_kVp =settings[6]
                                XRay_uA=settings[7]

                                if XRay_kVp.find("kVp")<0:
                                    XRay_kVp =settings[7]
                                    XRay_uA=settings[8]
                                    
                            else:
                                GainCorrect =False
                                DefectCorrect=False
                                XRay_kVp =settings[5]
                                XRay_uA=settings[6]

                            if XRay_kVp.find("kVp")>=0 and binning.find("Binx11"):
                                ### write to file: BuildLabel,Model,Serial,Binning,FW,XRaykVP,XRay_uA,ExpTime,GainCorrect,DefectCorrect,ImageAvg
                                f.write(theBuildLabel+","+theModel+","+theSerNo+","+\
                                        binning+","+FW+","+\
                                        XRay_kVp+","+XRay_uA+","+\
                                        expTime+","+str(GainCorrect)+","+str(DefectCorrect)+","+imgavg+"\n")

    f.close()
    
    return

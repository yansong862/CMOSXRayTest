import time
import copy
import winsound
import easygui
import os
import sys
import re
import ctypes
import collections
import shutil
import glob
import filecmp
import ntpath


import GUI_listOption as GetUserSelection


mswindows = (sys.platform == "win32")

_win_diskusage = collections.namedtuple('usage', 'total used free')



def pkiWarningBox(theMsg):
    easygui.msgbox(theMsg, "Warning...", 'OK', None, None)


def pkiInfoBox(theMsg):
    easygui.msgbox(theMsg, "Info...", 'OK', None, None)


def pkiYNBox(theMsg, theTitle):
    return easygui.ynbox(theMsg, theTitle, ('Yes', 'No'), None)    

def pkiWaitMessage(theMsg, intervalInS, intterValCnt):
    if theMsg != None:
        print(theMsg)

    count=0
    while(count<intterValCnt):
        print(".");
        time.sleep(intervalInS)
        count +=1

    return


    

### split the director and filename, return filename (wo directory)
def path_leaf(path):
    return path.strip('/').strip('\\').split('/')[-1].split('\\')[-1]


def getLatestDir(rootdir, filepattern ="*", bDebug =False):
    if bDebug: print "searching the latest dir under ", rootdir

    lastDir =""
    lastDirTime =0

    for filename in glob.glob(os.path.join(rootdir,filepattern)):
        #print "\n\n"
        #print filename

        (mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime) =os.stat(filename)

        if ctime>lastDirTime:
            lastDirTime =ctime
            lastDir =filename

    if bDebug: print "the last dir ", lastDir
    return lastDir

def getLatestFile(thedir, filepattern, bDebug =True):
    searchdir =os.path.join(thedir,filepattern)
    if bDebug: print (searchdir)
    bFoundFile =False
    lastFile =""
    lastFileTime =0
    
    for filename in glob.glob(searchdir):
        if (bDebug): print (filename)
        (mode, ino, dev, nlink, uid, gid, size, atime, mtime, ctime) =os.stat(filename)

        if ctime>lastFileTime:
            lastFileTime =ctime
            lastFile =filename
            bFoundFile =True

    if bDebug: print ("latest file: " +lastFile)
    return bFoundFile, lastFile



def doneTestBefore(dir_cmos_data, det_model, det_sn):
    print("\ndoneTestBefore?")
    xray_dir =os.path.join(dir_cmos_data, det_model +"-"+ det_sn)
    xray_dir =os.path.join(xray_dir, "X-Ray")
    print(xray_dir)

    if os.path.exists(xray_dir):
        print("found test record")
        lastXRayRecord =getLatestDir(xray_dir)
        print("latest record: " + lastXRayRecord)

        for filename in glob.glob(os.path.join(lastXRayRecord+"\\Test Record\\*Test Record*.xls")):
            print("\nFound previous test record: "+filename)
            return True
        
    print("found no x-ray record")
    return False


    

def getLinearityData_old(dir_report_root, dir_output_parent, detMod, detSN, bGigE, bSlowMode, rmaNo='None', bRMAUnit=False):
    try:
        rtnmsg ="Linearity Data is ready"

        ### clean up the previous linearity data
        dir_loc_linearity =os.path.join(dir_output_parent, 'Linearity')
        if os.path.exists(dir_loc_linearity):
            deleteDir(dir_loc_linearity)
        print "Removed", dir_loc_linearity
        
        lastSetupDir =getLatestDir(dir_report_root+detMod+'-'+detSN+'\\Setup\\')
        ### check the existence of the linearity data from light test.
        if bRMAUnit:
            lastSetupDir =dir_report_root+'\\'+detMod +'\\' + rmaNo+ '_' +detSN
        src =os.path.join(lastSetupDir, 'Linearity')
        print "checking linearity data dir", src

        ### is directory exist
        if not os.path.exists(src):
            #print src, "not found."
            rtnmsg ="Directory" + src + "not found"
            return False, rtnmsg

        ### is linearity data file from light test exist
        ''' not getting excel file from light test anymore
        pattern_xls ="Linearity Test_" +detSN +".xlsx"
        fcnt=0
        for filename in glob.glob(os.path.join(src, pattern_xls)):
            fcnt +=1
            print filename

        if not fcnt==1:
            #print "wrong number of files for searching of", pattern_xls
            rtnmsg = "error in finding linearity data from light test result" + pattern_xls
            return False, rtnmsg

        print "found ", pattern_xls
        '''

        ### is txt exist
        pattern_txt ="linearity.txt"
        fcnt=0
        for filename in glob.glob(os.path.join(src, pattern_txt)):
            fcnt +=1
            print filename

        if not fcnt==1:
            #print "wrong number of files for searching of", pattern_txt
            rtnmsg ="error in finding linearity data from light test result" +pattern_txt
            return False, rtnmsg

        print "found ", pattern_txt

        '''copy to local'''
        os.makedirs(dir_loc_linearity, mode =511)
        print "make new Linearity directory", dir_loc_linearity
        #shutil.copy(os.path.join(src, pattern_xls), dir_loc_linearity)
        #print "copy", pattern_xls, "to", dir_loc_linearity
        shutil.copy(os.path.join(src, pattern_txt), dir_loc_linearity)
        print "copy", pattern_txt, "to", dir_loc_linearity

	'''check the linearity data, and prepare data for 1-sensor only'''
	theCmd = "perl prepLinearityTxt.pl \""+ dir_loc_linearity +"\""

	getStatusOutput(theCmd)
		
        linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test_SerialNo.xlsx"

        if bGigE:
            linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test GigE.xlsx"
            
        if bSlowMode:
            linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test SlowMamo.xlsx"

        if "2923" in detMod and \
           bGigE and\
           bSlowMode:
            linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test GigE_Slow_2923.xlsx"
            
            
        if not os.path.exists(linearity_xls_template):
            rtnmsg = "not able to find " +linearity_xls_tempate
            return False, rtnmsg
        shutil.copy(linearity_xls_template, os.path.join(dir_loc_linearity, "LinearityData.xlsx"))
        print "copy", linearity_xls_template, "to", os.path.join(dir_loc_linearity, "LinearityData.xlsx")
            
        return True, rtnmsg
    except:
        return False, "Unknow error"


def getLinearityData(SCProdServer, RMAServer, UKCofCDir, dir_output_parent, detMod, detSN, bGigE, bSlowMode, rmaNo='None', bRMAUnit=False):
    try:
        bDebug =True

        bFoundLinearityData =False
        rtnmsg ="Searching Linearity Data ..."

        if os.path.exists(SCProdServer):
            linearParentDir =getLatestDir(SCProdServer+detMod+'-'+detSN+'\\Setup\\') # production server
            linearTxtDataDir =str(os.path.join(linearParentDir, 'Linearity'))
            linearXlDataDir =str(os.path.join(linearParentDir, 'Test Record'))            
            (bFoundLinearityData, rtnmsg) =getLinearityData_SCProd_RMA(linearXlDataDir, linearTxtDataDir, dir_output_parent, detMod, detSN, bGigE, bSlowMode)

            if (not bFoundLinearityData) and detSN.find('0')==0: #try the serial num without leading '0'
                linearParentDir =getLatestDir(SCProdServer+detMod+'-'+str(int(detSN))+'\\Setup\\') # production server
                linearTxtDataDir =str(os.path.join(linearParentDir, 'Linearity'))
                linearXlDataDir =str(os.path.join(linearParentDir, 'Test Record'))            
                (bFoundLinearityData, rtnmsg) =getLinearityData_SCProd_RMA(linearXlDataDir, linearTxtDataDir, dir_output_parent, detMod, detSN, bGigE, bSlowMode)
        else:
            print("\nskip " +SCProdServer)

        if (not bFoundLinearityData) and os.path.exists(RMAServer):
            linearParentDir =RMAServer+'\\'+detMod +'\\' + rmaNo+ '_' +detSN
            linearTxtDataDir =str(os.path.join(linearParentDir, 'Linearity'))
            linearXlDataDir =str(os.path.join(linearParentDir, 'X-Ray\\Test Record'))
            (bFoundLinearityData, rtnmsg) =getLinearityData_SCProd_RMA(linearXlDataDir, linearTxtDataDir, dir_output_parent, detMod, detSN, bGigE, bSlowMode)

            if (not bFoundLinearityData) and detSN.find('0')==0: #try the serial num without leading '0'
                linearParentDir =RMAServer+'\\'+detMod +'\\' + rmaNo+ '_' +str(int(detSN))
                linearTxtDataDir =str(os.path.join(linearParentDir, 'Linearity'))
                linearXlDataDir =str(os.path.join(linearParentDir, 'X-Ray\\Test Record'))
                (bFoundLinearityData, rtnmsg) =getLinearityData_SCProd_RMA(linearXlDataDir, linearTxtDataDir, dir_output_parent, detMod, detSN, bGigE, bSlowMode)
        else:
            print("\nskip " +RMAServer)
                
        if (not bFoundLinearityData) and os.path.exists(UKCofCDir):
            (bFoundLinearityData, rtnmsg) =getLinearityData_UKCofC(UKCofCDir, dir_output_parent, detMod, detSN)
            
            if (not bFoundLinearityData) and detSN.find('0')==0: #try the serial num without leading '0'
                (bFoundLinearityData, rtnmsg) =getLinearityData_UKCofC(UKCofCDir, dir_output_parent, detMod, detSN)
        else:
            print("\nskip " +UKCofCDir)
            

        return bFoundLinearityData, rtnmsg
    except:
        return False, "Unknow error"

    

def getLinearityData_SCProd_RMA(linearXlDataDir, linearTxtDataDir, dir_output_parent, detMod, detSN, bGigE, bSlowMode):
    try:
        bDebug =True
        
        rtnmsg ="\nSearching Linearity Data under the following directories:\n\t" + linearXlDataDir +"\n\t" +linearTxtDataDir
        print (rtnmsg)

        theLinearDataFile =""
        bFoundFile =False
        bTxtFile =False
        
        ### clean up the previous linearity data
        dir_loc_linearity =os.path.join(dir_output_parent, 'Linearity')
        if os.path.exists(dir_loc_linearity):
            deleteDir(dir_loc_linearity)
        print("Removed " + dir_loc_linearity)

        print (linearTxtDataDir)
        print (linearXlDataDir)
        ### is directory exist on production server?
        if not os.path.exists(linearTxtDataDir): #if linearTxtDataDir does not exist, linearXlDataDir won't exist either.
            rtnmsg ="Directory" + linearTxtDataDir +" or " + linearXlDataDir+ "not found"
            if bDebug: print(rtnmsg)
            return False, rtnmsg

        ### linearity in xl or txt?
        if os.path.exists(linearXlDataDir): #found "Test Record"
            filepattern ='Linearity*.xl*'

            (bFoundFile, theLinearDataFile) =getLatestFile(linearXlDataDir, filepattern)
            if bFoundFile:
                bTxtFile =False

        if (not bFoundFile) and os.path.exists(linearTxtDataDir): #not found linearity xl file, search for linearity txt data then
            filepattern ="linearity.txt"
            (bFoundFile, theLinearDataFile) =getLatestFile(linearTxtDataDir, filepattern)

            if bFoundFile:
                bTxtFile =True
            else:
                rtnmsg ="Directory" + linearTxtDataDir +" or " + linearXlDataDir+ "not found"
                return False, rtnmsg

        '''copy to local'''
        os.makedirs(dir_loc_linearity, mode =511)
        print "make new Linearity directory", dir_loc_linearity
        shutil.copy(theLinearDataFile, dir_loc_linearity)
        print "copy", theLinearDataFile, "to", dir_loc_linearity

        ### if txt data file, remove blank lines, copy the template
        if bTxtFile:
            '''check the linearity data, and prepare data for 1-sensor only'''
            theCmd = "perl prepLinearityTxt.pl \""+ dir_loc_linearity +"\""
            getStatusOutput(theCmd)

            linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test_SerialNo.xlsx"

            if bGigE:
                linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test GigE.xlsx"
                
            if bSlowMode:
                linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test SlowMamo.xlsx"
                
            if "2923" in detMod and \
               bGigE and\
               bSlowMode:
                linearity_xls_template ="C:\\CMOS\\ForReports\\"+"Linearity Test GigE_Slow_2923.xlsx"
                
                
            if not os.path.exists(linearity_xls_template):
                rtnmsg = "not able to find " +linearity_xls_tempate
                return False, rtnmsg
            shutil.copy(linearity_xls_template, os.path.join(dir_loc_linearity, "LinearityData.xlsx"))
            print "copy", linearity_xls_template, "to", os.path.join(dir_loc_linearity, "LinearityData.xlsx")
            
        return True, rtnmsg
    except:
        print ("Exception during linearity data serach")
        return False, "Unknow error"


def getLinearityData_UKCofC(dir_UKCofC, dir_output_parent, detMod, detSN):
    try:
        rtnmsg ="Searching Linearity Data under UKCofC dir..."
        print (rtnmsg)

        ### clean up the previous linearity data
        dir_loc_linearity =os.path.join(dir_output_parent, 'Linearity')
        if os.path.exists(dir_loc_linearity):
            deleteDir(dir_loc_linearity)
        print("Removed " + dir_loc_linearity)

        searchPattern =str(os.path.join(dir_UKCofC, str(detMod)+"*"+str(int(detSN))+"*Test Recrod*.xl*"))
        for filename in glob.glob(searchPattern):
            '''copy to local'''
            os.makedirs(dir_loc_linearity, mode =511)
            print "make new Linearity directory", dir_loc_linearity
            shutil.copy(filename, dir_loc_linearity)
            print "copy", filename, "to", dir_loc_linearity
            return True, "Found " +filename
                           
        return False, "Not found test record in " +dir_UKCofC
    except:
        print ("Exception during linearity data serach")
        return False, "Unknow error"


    


def uploadCofCLinearityData(srcdir,
                            destdir):
    print "upload linearity into CofC"

    retval =False
    PFResult="Done uploadCofCLinearityData"

    ### check linearity directory
    if not os.path.exists(srcdir):
        return False, srcdir + " Linearity directory doesn't exist"

    ### check test record directory
    if not os.path.exists(destdir):
        return False, destdir + " doesn't exist"

    ### check linearity data type
    filename =None
    filetype =None
    txtfilecnt =0
    for filename in glob.glob(os.path.join(srcdir, '*.txt')):
        txtfilecnt +=1
        filetype ='txt'
        break

    if txtfilecnt==0: #not found text data, assume that the excel test record will be use
        for filename in glob.glob(os.path.join(srcdir, '*.xls*')): #both xls and xlsx
            filename =os.path.basename(filename)
            filetype ='xl'
            break
        
    if filetype == None or filename ==None:
        return False, "Not found linearity data under directory "+srcdir
    
    if mswindows:
        print("filename " +filename)
        theCmd = "perl updateCofCLinearity.pl \""+ srcdir +"\" \"" + filename + "\" \"" + destdir + "\" " +filetype 

        result = getStatusOutput(theCmd)
        retval =True
        
    else:
        print "not support this OS"
    
    return retval, PFResult


def updateCofCTemplate(destdir,
                       temp_csv_output,
                       buildLabel,
                       buildRev,
                       ADCFirmRev,
                       DAQFirmRev,
                       bSlowMode,
                       bMTFHighResolution,
                       bNewGigENoiseSpec,
                       bRMA):

    print ("\nupdate CofC template")

    retval =False
    PFResult="Done updateCofCTemplate"

    ### check test record directory
    if not os.path.exists(destdir):
        return False, destdir + " doesn't exist. No test record found."

    if mswindows:
        theCmd = "perl updateCofCTemplate.pl \""+ destdir + "\" \"" \
                 + temp_csv_output +"\" "\
                 + buildLabel +" "\
                 + buildRev +" "\
                 + ADCFirmRev +" "\
                 + DAQFirmRev +" "\
                 + str(bSlowMode)+ " "\
                 + str(bMTFHighResolution)+ " " \
                 + str(bNewGigENoiseSpec) + " "\
                 + str(bRMA)

        result = getStatusOutput(theCmd)
        retval =True
        
    else:
        print "not support this OS"
    
    return retval, PFResult


def printCofC(destdir):

    print ("\nprint CofC")

    retval =False
    PFResult="Done print CofC"

    ### check test record directory
    if not os.path.exists(destdir):
        return False, destdir + " doesn't exist. No test record found."

    if mswindows:
        theCmd = "perl printCofC.pl \""+ destdir + "\""

        result = getStatusOutput(theCmd)
        retval =True
        
    else:
        print "not support this OS"
    
    return retval, PFResult

def getCofCTestResult(temp_csv_output):
    print "getCofCTestResult"

    retval =False
    PFResult="UNKNOW"

    if os.path.exists(temp_csv_output):
        with open(temp_csv_output) as f:
            for line in f:
                if bool('Overall Result' in line):
                    if bool('PASS' in line):
                        PFResult="PASS"
                        retval =True
                        print("Passed CMOS X-Ray test.")
                    else:
                        PFResult="FAIL"
                        print("Failed CMOS X-Ray test.")
                        retval =False

    else:
        print(temp_csv_output+ " doesn't exist")

    return retval, PFResult


def updateCofCLinearityData(srcdir,
                            destdir,
                            temp_csv_output,
                            buildLabel,
                            buildRev,
                            ADCFirmRev,
                            DAQFirmRev,
                            bSlowMode,
                            bMTFHighResolution,
                            bNewGigENoiseSpec,
                            bRMA):
    retval =False
    PFResult="UNKNOW"
              
    (retval, PFResult) =uploadCofCLinearityData(srcdir, destdir)

    if retval:
        (retval, PFResult) =updateCofCTemplate(destdir,
                                               temp_csv_output,
                                               buildLabel,
                                               buildRev,
                                               ADCFirmRev,
                                               DAQFirmRev,
                                               bSlowMode,
                                               bMTFHighResolution,
                                               bNewGigENoiseSpec,
                                               bRMA)

    if retval:
        print("True")
    else:
        print("False")
    
    if retval:
        (retval, PFResult) =getCofCTestResult(temp_csv_output)
        
              
              

    '''
    print "update linearity in CofC", srcdir,destdir

    retval =False
    PFResult="UNKNOW"

    if not os.path.exists(srcdir):
        return False, srcdir + " Linearity directory doesn't exist"

    ### check linearity data type
    filename =None
    filetype =None
    txtfilecnt =0
    for filename in glob.glob(os.path.join(src, '*.txt')):
        txtfilecnt +=1
        filetype ='txt'
        break

    if txtfilecnt==0: #not found text data, assume that the excel test record will be use
        for filename in glob.glob(os.path.join(src, '*.xls*')): #both xls and xlsx
            filetype ='xl'
            break
        
    if filetype == None or filename ==None:
        return False, "Not found linearity data under directory "+srcdir
    
    if mswindows:
        #old theCmd = "perl updateLinearityData.pl \""+ srcdir +"\" \"" + destdir + "\"" + " " +temp_csv_output + " " + buildLabel+ " " +str(bSlowMode)+ " " +str(bMTFHighResolution)+ " " +str(bNewGigENoiseSpec)
        theCmd = "perl updateCofCLinearity.pl \""+ srcdir +"\" \"" + destdir + "\" " + filename + " " +filetype

        result = getStatusOutput(theCmd)

        theCmd = "perl updateCofCTemplate.pl \""+ destdir + "\" " \
                 + temp_csv_output + " "  \
                 + buildLabel +" "\
                 + buildRev +" "\
                 + ADCFirmRev +" "\
                 + DAQFirmRev +" "\
                 + buildLabel +" "\
                 + str(bSlowMode)+ " "\
                 + str(bMTFHighResolution)+ " " \
                 + str(bNewGigENoiseSpec)


        result = getStatusOutput(theCmd)

        
        if os.path.exists(temp_csv_output):
            with open(temp_csv_output) as f:
                for line in f:
                    if bool('Overall Result' in line):
                        if bool('PASS' in line):
                            PFResult="PASS"
                        else:
                            PFResult="FAIL"

    else:
        print "not support this OS"
    '''
    
    return retval, PFResult
    





def getADCBoardInfo(dir_cmosReport, dir_xrayReportApp, bGrade=True):
    filename ='C:/CMOS/Configs/CMOSDetectorADCList.txt'
    print filename


    modelBuildList=[]
    buildRevList=[]
    ModelBuildADCList=[]
    ADCFirmRevList=[]
    DAQFirmRevList =[]
    GigEList=[]
    SlowModeList=[]
    MTFResolutionList=[]
    NoiseSpecList=[]
    with open(filename, "r") as ins:
        for line in ins:
            if line.find("#")<0: #not comment line
                line =line.strip()
                line =line.replace(" ","")
                line =line.replace("\t","")
                line =re.split('[;|\n]',line)
                if len(line)>2:
                    print "line: ", line, len(line)
                    modelBuildList.append(line[0])
                    buildRevList.append(line[1])
                    ModelBuildADCList.append(line[2])
                    ADCFirmRevList.append(line[3])
                    DAQFirmRevList.append(line[4])
                    GigEList.append(line[5])
                    SlowModeList.append(line[6])
                    MTFResolutionList.append(line[7])
                    NoiseSpecList.append(line[8])

    modelBuildList.append("Not found")
    buildRevList.append("Not found")
    ModelBuildADCList.append("None")
    ADCFirmRevList.append("None")
    DAQFirmRevList.append("None")
    GigEList.append("None")
    SlowModeList.append("None")
    MTFResolutionList.append("None")
    NoiseSpecList.append("None")
    
    if (len(modelBuildList) != len(ModelBuildADCList)):
        print filename, " has error."
        sys.exit()
        
    print "\n modelBuildList:\n",modelBuildList, len(modelBuildList)
    print "\n ModelBuildADCList:\n",ModelBuildADCList, len(ModelBuildADCList)


    title ="Select one build"
    maxNumWords =len(max(modelBuildList,key=len))
    maxNumWords =max(maxNumWords, len(title)+20)


    GUI_options = {'title' : title, 'selectmode' : 'single', 'width':str(maxNumWords+5), 'height':str(len(modelBuildList)) }
    thebuild =GetUserSelection.easyListBox(modelBuildList, **GUI_options)
    buildSel =thebuild.getInput()

    while not easygui.ynbox('Please confirm model No.: "' + buildSel+'"',
                             'Detector model info:', ('Yes', 'No'), None):
        thebuild =GetUserSelection.easyListBox(modelBuildList, **GUI_options)
        buildSel =thebuild.getInput()        

    theBuildIndex =modelBuildList.index(buildSel)
    theBuildRev =buildRevList[theBuildIndex]
    theADCBoard =ModelBuildADCList[theBuildIndex]
    theADCFirmRev=ADCFirmRevList[theBuildIndex]
    theDAQFirmRev =DAQFirmRevList[theBuildIndex]
    theGigE =GigEList[theBuildIndex]
    theSlowMode =SlowModeList[theBuildIndex]
    theMTFResolution =MTFResolutionList[theBuildIndex]
    theNoiseSpec =NoiseSpecList[theBuildIndex]
    
    print "\n index of ", str(buildSel), "is", theBuildIndex, ";", theADCBoard, "is selected."


    if bGrade:
        ### copy the corresponding spec
        src_spec =os.path.join(dir_cmosReport, "DefectGrades_Standard.txt")
        dest_spec =os.path.join(dir_xrayReportApp, "DefectGrades.txt")
        if "1512-C-3D-C600B" in buildSel:
            src_spec =os.path.join(dir_cmosReport, "DefectGrades_1512-C-3D-C600B.txt")
        #elif "1207N-C16-GADM" in buildSel:
        #    src_spec =os.path.join(dir_cmosReport, "DefectGrades_Mammo.txt")
        elif "2923N-C22-HRCC" in buildSel:
            src_spec =os.path.join(dir_cmosReport, "DefectGrades_Gold.txt")
            
            
        shutil.copy(src_spec, dest_spec)
        print("copy " +src_spec +" ==> " +dest_spec)

    return str(buildSel), theBuildRev, theADCBoard, theADCFirmRev, theDAQFirmRev, theGigE, theSlowMode,theMTFResolution,theNoiseSpec



def getModelADCBoardInfo(sModel, dir_cmosReport, dir_xrayReportApp, bGrade=True):
    filename ='C:/CMOS/Configs/CMOSDetectorADCList.txt'
    print filename


    modelBuildList=[]
    buildRevList=[]
    ModelBuildADCList=[]
    ADCFirmRevList=[]
    DAQFirmRevList =[]
    GigEList=[]
    SlowModeList=[]
    MTFResolutionList=[]
    NoiseSpecList=[]
    with open(filename, "r") as ins:
        for line in ins:
            if line.find("#")<0: #not comment line
                line =line.strip()
                line =line.replace(" ","")
                line =line.replace("\t","")
                line =re.split('[;|\n]',line)
                if len(line)>2 and \
                   sModel in line[0]:
                    print "line: ", line, len(line)
                    modelBuildList.append(line[0])
                    buildRevList.append(line[1])
                    ModelBuildADCList.append(line[2])
                    ADCFirmRevList.append(line[3])
                    DAQFirmRevList.append(line[4])
                    GigEList.append(line[5])
                    SlowModeList.append(line[6])
                    MTFResolutionList.append(line[7])
                    NoiseSpecList.append(line[8])

    modelBuildList.append("Not found")
    buildRevList.append("Not found")
    ModelBuildADCList.append("None")
    ADCFirmRevList.append("None")
    DAQFirmRevList.append("None")
    GigEList.append("None")
    SlowModeList.append("None")
    MTFResolutionList.append("None")
    NoiseSpecList.append("None")
    
    if (len(modelBuildList) != len(ModelBuildADCList)):
        print filename, " has error."
        sys.exit()
        
    print "\n modelBuildList:\n",modelBuildList, len(modelBuildList)
    print "\n ModelBuildADCList:\n",ModelBuildADCList, len(ModelBuildADCList)


    title ="Select one build"
    maxNumWords =len(max(modelBuildList,key=len))
    maxNumWords =max(maxNumWords, len(title)+20)


    GUI_options = {'title' : title, 'selectmode' : 'single', 'width':str(maxNumWords+5), 'height':str(len(modelBuildList)) }
    thebuild =GetUserSelection.easyListBox(modelBuildList, **GUI_options)
    buildSel =thebuild.getInput()

    while not easygui.ynbox('Please confirm model No.: "' + buildSel+'"',
                             'Detector model info:', ('Yes', 'No'), None):
        thebuild =GetUserSelection.easyListBox(modelBuildList, **GUI_options)
        buildSel =thebuild.getInput()        

    theBuildIndex =modelBuildList.index(buildSel)
    theBuildRev =buildRevList[theBuildIndex]
    theADCBoard =ModelBuildADCList[theBuildIndex]
    theADCFirmRev=ADCFirmRevList[theBuildIndex]
    theDAQFirmRev =DAQFirmRevList[theBuildIndex]
    theGigE =GigEList[theBuildIndex]
    theSlowMode =SlowModeList[theBuildIndex]
    theMTFResolution =MTFResolutionList[theBuildIndex]
    theNoiseSpec =NoiseSpecList[theBuildIndex]
    
    print "\n index of ", str(buildSel), "is", theBuildIndex, ";", theADCBoard, "is selected."


    if bGrade:
        ### copy the corresponding spec
        src_spec =os.path.join(dir_cmosReport, "DefectGrades_Standard.txt")
        dest_spec =os.path.join(dir_xrayReportApp, "DefectGrades.txt")
        if "1512-C-3D-C600B" in buildSel:
            src_spec =os.path.join(dir_cmosReport, "DefectGrades_1512-C-3D-C600B.txt")
        #elif "1207N-C16-GADM" in buildSel:
        #    src_spec =os.path.join(dir_cmosReport, "DefectGrades_Mammo.txt")
        elif "2923N-C22-HRCC" in buildSel:
            src_spec =os.path.join(dir_cmosReport, "DefectGrades_Gold.txt")
            
        shutil.copy(src_spec, dest_spec)
        print("copy " +src_spec +" ==> " +dest_spec)

    return str(buildSel), theBuildRev, theADCBoard, theADCFirmRev, theDAQFirmRev, theGigE, theSlowMode,theMTFResolution,theNoiseSpec





def loadReportToDB(mod_no, serial_no, rma_no, buildLabel, outputdir):
    print "load report to DB";
    if mswindows: 
        theCmd = "perl grade.pl "+ mod_no +" " + serial_no +" " + rma_no+" " + buildLabel+ " \"" + outputdir + "\""
    else:
        print "not support this OS"

    result = getStatusOutput(theCmd)

    print "resutl of load DB", result[0]
    return

    
def getStatusOutput(theCmd):
    """Return (status, output) of executing cmd in a shell."""
    print("run command " + theCmd)
    if not mswindows:
        return commands.getStatusOutput(theCmd)
    pipe = os.popen(theCmd + ' 2>&1', 'r')
    text = pipe.read()
    sts = pipe.close()
    if sts is None: sts = 0
    if text[-1:] == '\n': text = text[:-1]
    print("return of the command exe: " +text)
    return sts, text


def deleteDir(path):
    """deletes the path entirely"""
    if mswindows: 
        theCmd = "RMDIR "+ path +" /s /q"
    else:
        theCmd = "rm -rf "+path
    result = getStatusOutput(theCmd)
    if(result[0]!=0):
        raise RuntimeError(result[1])
    
def prepDir(thePath):
    if os.path.exists(thePath):
        print "remove output directory", thePath
        deleteDir(thePath)

    os.makedirs(thePath, mode=511)
    

def isDataReadyForDelivery(dir_delivery):
    print "checking files in ", dir_delivery
    if not os.path.exists(dir_delivery):
        return False
    return True




def copyImages2Delivery(dir_image_src, dir_AutoCal, dir_Darks, dir_Floods, dir_DefectMap,\
                        dir_custDelivery, dir_TestImages, dir_SupportFiles):
    print("\n\n\ncopyImages2Delivery")
    if not os.path.exists(dir_image_src):
        print(dir_image_src +" doesn't exist.")
        return False
    if not os.path.exists(dir_AutoCal):
        print(dir_AutoCal +" doesn't exist.")
        return False
    if not os.path.exists(dir_Darks):
        print(dir_Darks +" doesn't exist.")
        return False
    if not os.path.exists(dir_Floods):
        print(dir_Floods +" doesn't exist.")
        return False
    if not os.path.exists(dir_DefectMap):
        print(dir_DefectMap +" doesn't exist.")
        return False


    ### re-create customer delivery folder:
    if os.path.exists(dir_custDelivery):
        deleteDir("\""+dir_custDelivery+"\"")

    os.makedirs(dir_custDelivery, mode=511)
    if not os.path.exists(dir_custDelivery):
        print( dir_custDelivery + " cannot be created")
        return False

    os.makedirs(dir_TestImages, mode=511)
    if not os.path.exists(dir_TestImages):
        print(dir_TestImages + " cannot be created")
        return False

    os.makedirs(dir_SupportFiles, mode=511)
    if not os.path.exists(dir_SupportFiles):
        print(dir_SupportFiles + " cannot be created")
        return False


    bCopiedAll =True
    copyCnt =0
    ### copy image files to dir_TestImages
    for filename in glob.glob(os.path.join(dir_image_src, '*DefectCorrected*hand*phantom*.tif')):
        print("copy " + filename + " from " + dir_image_src + " to " + dir_TestImages)
        shutil.copy(filename, dir_TestImages)
        copyCnt =copyCnt+1

    if copyCnt<2:
        print("only copied "+str(copyCnt)+ "hand phantom images. The minimum should be 2")
        bCopiedAll =False

    
    copyCnt =0
    for filename in glob.glob(os.path.join(dir_image_src, '*DefectCorrected*digimam*phantom*.tif')):
        print("copy " + filename + " from " + dir_image_src + " to " + dir_TestImages)
        shutil.copy(filename, dir_TestImages)
        copyCnt =copyCnt+1

    if copyCnt<1:
        print("only copied "+str(copyCnt)+ "digimam phantom image. The minimum should be 1")
        bCopiedAll =False

    copyCnt =0
    for filename in glob.glob(os.path.join(dir_image_src, '*DefectCorrected*CIRS*phantom*.tif')):
        print("copy " + filename + " from " + dir_image_src + " to " + dir_TestImages)
        shutil.copy(filename, dir_TestImages)
        copyCnt =copyCnt+1

    if copyCnt<1:
        print("only copied "+str(copyCnt)+ "CIRS phantom image. The minimum should be 1")
        bCopiedAll =False

    copyCnt =0
    for filename in glob.glob(os.path.join(dir_image_src, '*DefectCorrected*TOR*MAS*phantom*.tif')):
        print("copy " + filename + " from " + dir_image_src + " to " + dir_TestImages)
        shutil.copy(filename, dir_TestImages)
        copyCnt =copyCnt+1

    if copyCnt<1:
        print("only copied "+str(copyCnt)+ "TOR-MAS phantom image. The minimum should be 1")
        bCopiedAll =False

    copyCnt =0
    for filename in glob.glob(os.path.join(dir_image_src, '*DefectCorrected*MTF*phantom*.tif')):
        print("copy " + filename + " from " + dir_image_src + " to " + dir_TestImages)
        shutil.copy(filename, dir_TestImages)
        copyCnt =copyCnt+1

    if copyCnt<3:
        print("only copied "+str(copyCnt)+ "MTF phantom images. The minimum should be 3")
        bCopiedAll =False


    ### copy support files to dir_SupportFiles
    copyCnt =0
    for filename in glob.glob(os.path.join(dir_AutoCal, '*autocal*.txt')):
        print("copy " + filename + " from " + dir_AutoCal + " to " + dir_SupportFiles)
        shutil.copy(filename, dir_SupportFiles)
        copyCnt =copyCnt+1

    if copyCnt<1:
        print("only copied "+str(copyCnt)+ "autocal.txt. The minimum should be 1")
        bCopiedAll =False

    copyCnt =0
    for filename in glob.glob(os.path.join(dir_Darks, '*dark.smv')):
        print("copy " + filename + " from " + dir_Darks + " to " + dir_SupportFiles)
        shutil.copy(filename, dir_SupportFiles)
        copyCnt =copyCnt+1

    if copyCnt<2:
        print("only copied "+str(copyCnt)+ "darks.smv. The minimum should be 2")
        bCopiedAll =False
    

    copyCnt =0
    for filename in glob.glob(os.path.join(dir_Floods, '*flood.smv')):
        print("copy " + filename + " from " + dir_Floods + " to " + dir_SupportFiles)
        shutil.copy(filename, dir_SupportFiles)
        copyCnt =copyCnt+1

    if copyCnt<2:
        print("only copied "+str(copyCnt)+ "floods.smv. The minimum should be 2")
        bCopiedAll =False

    copyCnt =0
    for filename in glob.glob(os.path.join(dir_DefectMap, '*_Defect_Map_*.smv')):
        print("copy " + filename + " from " + dir_DefectMap + " to " + dir_SupportFiles)
        shutil.copy(filename, dir_SupportFiles)
        copyCnt =copyCnt+1

    if copyCnt<2:
        print("only copied "+str(copyCnt)+ "_Defect_Map_*.smv. The minimum should be 2")
        bCopiedAll =False

    return bCopiedAll







def formatDisk_osPOpen(theDrv,devSN):
    print "formating", theDrv, "It will take a couple minutes. Please wait..."
    if mswindows: 
        theCmd = "echo Y | format "+ theDrv +" /FS:UDF"
    else:
        print "not support this OS"

    result = getStatusOutput(theCmd)
 

def formatDisk(theDrv,devSN):
    exeCommand = "echo Y | format "+ "D:" +" /FS:UDF"
    errCode =os.system(exeCommand)

    if errCode==0:
        return True;
    else:
        pkiWarningBox("Failed format CD. CD might not be created correctly")
        return False
 


      
def disk_usage(path):
    try:
        if not mswindows:
            raise NotImplementedError("platform not supported")
            return
            
        _, total, free = ctypes.c_ulonglong(), ctypes.c_ulonglong(), \
                           ctypes.c_ulonglong()
        if sys.version_info >= (3,) or isinstance(path, unicode):
            fun = ctypes.windll.kernel32.GetDiskFreeSpaceExW
        else:
            fun = ctypes.windll.kernel32.GetDiskFreeSpaceExA
        ret = fun(path, ctypes.byref(_), ctypes.byref(total), ctypes.byref(free))
        if ret == 0:
            raise ctypes.WinError()
        used = total.value - free.value
        return _win_diskusage(total.value, used, free.value)
    except Exception:
        print "Failed to get disk usage."
        return _win_diskusage(-1, -1, -1)

def bytes2human(n):
    symbols = ('K', 'M', 'G', 'T', 'P', 'E', 'Z', 'Y')
    prefix = {}
    for i, s in enumerate(symbols):
        prefix[s] = 1 << (i+1)*10
    for s in reversed(symbols):
        if n >= prefix[s]:
            value = float(n) / prefix[s]
            return '%.1f%s' % (value, s)
    return "%sB" % n


    
def isDiskReady(drv_dvd, title):
    print "checking Disk drive ", drv_dvd

    #rtnVal =formatDisk(drv_dvd, title)
    rtnVal =True
    formatDisk_osPOpen(drv_dvd, title)
    
    '''is any file on disk'''
    usage =disk_usage(drv_dvd)
    print "total", bytes2human(usage.total), ", free",  bytes2human(usage.free), ", used", bytes2human(usage.used)

    return rtnVal


def recursive_overwrite(src, dest, ignore=None):
    if os.path.isdir(src):
        print src
        if not os.path.isdir(dest):
            os.makedirs(dest)
        files = os.listdir(src)
        if ignore is not None:
            ignored = ignore(src, files)
        else:
            ignored = set()
        for f in files:
            if f not in ignored:
                recursive_overwrite(os.path.join(src, f), 
                                    os.path.join(dest, f), 
                                    ignore)
    else:
        print "copy from", src, "==>", dest
        shutil.copyfile(src, dest)


def data2Dvd(src, drv_dvd):
    print "copying data from ", src, " ==> ", drv_dvd, "It will take a couple minutes. Please wait..."
    if mswindows: 
        theCmd = "xcopy /E /V /Y \""+ src +"\" " + drv_dvd
    else:
        print "not support this OS"

    result = getStatusOutput(theCmd)
    return

    print "copying data from ", src, " ==> ", drv_dvd
    recursive_overwrite(src, drv_dvd)
    return True
    src_files = os.listdir(src)
    for file_name in src_files:
        full_file_name = os.path.join(src, file_name)
        if (os.path.isdir(full_file_name)):
            #os.makedirs(file_name,511)
            print full_file_name
            print "copying...." , file_name
        elif (os.path.isfile(full_file_name)):
            print full_file_name
            print "copying...." , file_name
            #shutil.copy(file_name, drv_dvd)
    
    #return True
    return sameFiles(src, drv_dvd)


def checkDataOnDvd(drv_dvd):
    print(str(len(glob.glob(drv_dvd))))
    print("check " +drv_dvd)
    if not os.path.isdir(drv_dvd+"\\"):
        print(drv_dvd +" drive doesn't exist")
        #return False

    if not os.path.isdir(os.path.join(drv_dvd,"Support Files")):
        print("not found folder: Support Files")
        return False
    else:        
        #ifiles =len(glob.glob(drv_dvd+"\\Support Files"))
        ifiles =len(glob.glob("D:\\Support Files"))
        if ifiles!=15:
            print("expect 16 files under 'Support Files'. actual found " +str(ifiles))
    
    
    if not os.path.isdir(drv_dvd+"\\Test Images"):
        print("not found folder: Test Images")
        return False
    else:
        ifiles =len(glob.glob(drv_dvd+"\\Test Images"))
        if ifiles!=11:
            print("expect 12 files under 'Test Images'. actual found " +str(ifiles))
    
    return True



def data2Server(src, dest, det_model, det_sn, bVfy=True):
    print "copying data from ", src, " ==> ", dest
    xray_dir =os.path.join(dest, det_model +"-"+ det_sn)
    xray_dir =os.path.join(xray_dir, "X-Ray")

    fcnt=0
    if os.path.exists(xray_dir):
        for filename in os.listdir(xray_dir):
            fcnt +=1
        print fcnt, "files found in directory", xray_dir
    else:
        os.makedirs(xray_dir)

    if fcnt==0:
        xray_dir =os.path.join(xray_dir, det_model +"-"+ det_sn)
        os.makedirs(xray_dir)
    else:
        xray_dir =os.path.join(xray_dir, det_model +"-"+ det_sn+"."+str(fcnt))
        os.makedirs(xray_dir)
    print "xray dir", xray_dir

    if mswindows: 
        theCmd = "xcopy /E /V /Y \""+ src +"\" " + xray_dir
    else:
        print "not support this OS"

    result = getStatusOutput(theCmd)

    if bVfy:
        return sameFiles(src, xray_dir)
    else:
        return True

def data2EngServer(src, dest, subdir, incr, det_model, det_sn):
    print "copying data from ", src, " ==> ", dest

    if (not os.path.exists(src)) or not (os.path.isdir(src)):
        print src, " directory doesn't exist";
        return False

    if (not os.path.exists(dest)) or not (os.path.isdir(dest)):
        print dest, " directory doesn't exist";
        return False


    xray_dir =os.path.join(dest, det_model +"-"+ det_sn)

    fcnt=0
    if os.path.exists(xray_dir):
        for filename in os.listdir(xray_dir):
            fcnt +=1

        if not bool(incr):
            fcnt -=1;
        print fcnt, "files found in directory", xray_dir
    else:
        os.makedirs(xray_dir)

    if fcnt==0:
        xray_dir =os.path.join(xray_dir, det_model +"-"+ det_sn)
    else:
        xray_dir =os.path.join(xray_dir, det_model +"-"+ det_sn+"."+str(fcnt))

    xray_dir =os.path.join(xray_dir, subdir)
    os.makedirs(xray_dir)
    
    print "xray dir", xray_dir

    if mswindows: 
        theCmd = "xcopy /E /V /Y \""+ src +"\" " + xray_dir
    else:
        print "not support this OS"

    result = getStatusOutput(theCmd)

    
    #return True
    return sameFiles(src, xray_dir)

def rmaData2Server(src, dest, det_model, det_sn, rmaNo, bInitTest):
    init_dir ='Initial Test'
    
    print "copying data from ", src, " ==> ", dest
    rma_dir =os.path.join(dest, det_model+"\\"+rmaNo +"_"+ det_sn)

    fcnt=0
    if os.path.exists(rma_dir):
        for filename in os.listdir(rma_dir):
            if 'X-Ray' in filename:
                fcnt +=1
        print fcnt, "X-Ray directory found in directory", rma_dir
    else:
        if bInitTest:
            os.makedirs(rma_dir)
        else:
            print("Not found RMA directory:", rma_dir)
            pkiWarningBox("RMA directory " + rma_dir + " not found. Data is not copied to server. Get engineer help")
            return False

    rma_xray_dir =os.path.join(rma_dir, 'X-Ray')
    if bInitTest:
        rma_xray_dir =os.path.join(rma_dir, init_dir)

        #orgdir =os.getcwd()
        #os.chdir(rma_dir)

        if os.path.exists(rma_xray_dir):
            print ("delete dir %s" %rma_xray_dir)
            deleteDir("\""+rma_xray_dir+"\"") ### deleted "Init Test" directory, if exists.

        print("create dir %s" %rma_xray_dir)
        #os.makedirs("\""+rma_xray_dir+"\"", mode =511)
        os.makedirs("%s" %rma_xray_dir,  mode =511)
        
        #os.chdir(orgdir)
        #print("current dir %s" %os.getcwd())
    else:
        if fcnt>0:
            rma_xray_dir =os.path.join(rma_dir, "X-Ray."+str(fcnt))

        os.makedirs("%s" %rma_xray_dir,  mode =511)

    print(rma_xray_dir)
    theCmd = "xcopy /E /V /Y \""+ src +"\" \"" + rma_xray_dir + "\""
    if mswindows: 
        theCmd = "xcopy /E /V /Y \""+ src +"\" \"" + rma_xray_dir + "\""
    else:
        print "not support this OS"

    result = getStatusOutput(theCmd)

    return sameFiles(src, rma_xray_dir)



def pkiCopyFiles(src, dest):
    if os.path.exists(src) and os.path.exists(dest):
        print("\nCopying from " + src + " to " +dest)
        if mswindows: 
            theCmd = "xcopy /E /V /Y \""+ src +"\" " + dest
        else:
            print "not support this OS"

        result = getStatusOutput(theCmd)
    else:
        print ("Trouble finding the following directories. Counldn't copy the filese.\n\t"+src+"\n\t"+dest)
        return False
    
    return True

    

def getLastXRayDataToLocal(serverDir, det_model, det_sn, rmaNo, bRMA):
    srcroot =serverDir
    dest ="c:\\testdata\\Hologic_0108_org1\\"
    if not bRMA:
        srcroot =os.path.join(serverDir, det_model +"-"+ det_sn+"\\X-Ray")

        if os.path.exists(srcroot):
            latesdir = getLatestDir(srcroot)
            listdir =os.listdir(srcroot)
            for thedir in listdir:
                if latesdir in os.path.join(srcroot,thedir):
                    pkiCopyFiles(latesdir, os.path.join(dest, thedir)+"\\")
            
    else:
        srcroot =os.path.join(serverDir, det_model+"\\"+rmaNo+"_"+det_sn)

    
    
    return True




def getOfflineImages(SCProdServer, RMAServer, dir_output_parent, detMod, detSN, rmaNo='None', bRMAUnit=False):
    try:
        bDebug =True

        bFoundImageData =False
        theXRayTestDir =os.path.join(SCProdServer,detMod+'-'+detSN,'X-Ray')
        if bDebug: print("Search image data under: " +theXRayTestDir)
        if os.path.exists(theXRayTestDir):
            xrayDataDir =getLatestDir(theXRayTestDir) # production server
            if bDebug: print("copy from: " +xrayDataDir)
            
            if pkiCopyFiles(xrayDataDir, dir_output_parent):
                return True

        theXRayTestDir =os.path.join(RMAServer, detMode, rmaNo+'_'+detSN)
        if bDebug: print("Search image data under: " +theXRayTestDir)
        if os.path.exists(theXRayTestDir):
            xrayDataDir =getLatestDir(theXRayTestDir, 'X-Ray')
            if bDebug: print("copy from: " +xrayDataDir)
            
            if pkiCopyFiles(xrayDataDir, dir_output_parent):
                return True

        theXRayTestDir =os.path.join(RMAServer, detMode, rmaNo+'_'+str(int(detSN)))
        if bDebug: print("Search image data under: " +theXRayTestDir)
        if os.path.exists(theXRayTestDir):
            xrayDataDir =getLatestDir(theXRayTestDir, 'X-Ray')
            if bDebug: print("copy from: " +xrayDataDir)
            
            if pkiCopyFiles(xrayDataDir, dir_output_parent):
                return True
        
        return False
    except:
        return False




### Returns True if recursively identical, False otherwise
def sameFiles(dir1, dir2):
    print "\n\n compare ", dir1, dir2
    comparison = filecmp.dircmp(dir1, dir2)
    numLOnly =len(comparison.left_only)
    numROnly =len(comparison.right_only)
    numDiff =len(comparison.diff_files)
    numFunny =len(comparison.funny_files)
    print "Found Left Only:", numLOnly
    print "Found Right Only:", numROnly
    print "Found Diff Files:", numDiff
    #print "Funny:", numFunny
    if (numLOnly + numROnly + numDiff +numFunny)>0:
        print "!!!found different files in ", dir1, dir2
        return False
    else:
        same_so_far = True
        for i in comparison.common_dirs:
            same_so_far = same_so_far and sameFiles(os.path.join(dir1, i), os.path.join(dir2, i))
            if not same_so_far:
                break

        if same_so_far:
            print "no different file found in ", dir1, dir2
        else:
            print "!!!found different files in ", dir1, dir2
            
        return same_so_far
    




def matchDetectData(filedir, det_mod, ser_no):
    fcnt=0
    for filename in glob.glob(os.path.join(filedir+"\\*"+det_mod+"*"+ser_no+"*")):
        fcnt+=1
        print(filename)

    print("found "+str(fcnt)+ " files.")
    if fcnt>1:
        return True;
    else:
        return False;
    

def renameTestRecords(recDir, buildLabel, serNo):
    if not os.path.exists (recDir):
        return

    for filename in glob.glob(os.path.join(recDir,"*")):
        istart =os.path.basename(filename).find(" Test Record")
        newname =buildLabel+"_"+serNo+os.path.basename(filename)[istart:len(os.path.basename(filename))]
	theCmd = "rename \""+ filename +"\" \""+newname+"\""

	getStatusOutput(theCmd)
	
    return

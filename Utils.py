import DexelaPy
import numpy
import time   
import math
import os.path

def SaveConfigFile(offsets,filename,model):
    header = "##########################################\n#\n#\n#\n#  (Use # at start of line for a comment line)\n#\n#\n#\n##########################\n"
    header +="####### Toggle CL Power##########################\n#\nWC0\n#\nWC1\n#\n#\n###########\n#Board 1\n###########\n"
    
    f = open(filename,'w+')
    f.write(header)
    
    if model != 1313:
        for n in range(len(offsets)):
            sensor = int(n/6+1)
            strip = int(n%6 +1)
            val = int(offsets[n])
            
            command = "W" + str(sensor) + "02" + str(strip) + str(val).zfill(5) + "\n"
            f.write(command)
    
    
    f.close()
    

def ConfigureGlobals():
    global model
    global sensorStrips
    global sensorX
    global sensorY
    global numSensors
    global sensorsH
    global sensorsV
    
    if model != 1313: 
        sensorStrips = 6
        sensorX = 1536
        sensorY = 1944
    
        if model == 2923 or model == 2321:
            numSensors = 4
            sensorsH = 2
            sensorsV = 2
    
        if model == 2315 or model == 2307:
            numSensors = 2
            sensorsH = 2
            sensorsV = 1
    
        if model == 1512 or model == 1207:
            numSensors = 1
            lookup = [i for i in range(6)]
            sensorsH = 1
            sensorsV = 1
        
        if model == 1207 or model == 2307:
            sensorY = 864
    else:
        print "Utils.py->ConfigureGlobals not support model 1313. Test will be stopped here!"
        sys.exit()

    print "model: ", model
    print "sensorStrips: ", sensorStrips
    print "sensorX: ", sensorX
    print "sensorY: ", sensorY
    print "numSensors: ", numSensors
    print "sensorsH: ", sensorsH
    print "sensorsV: ", sensorsV
    
    
    
    
def sqrDiff(img,avg,x1,y1,x2,y2,z):
    avg2 = 0.0
    counter = 0
    val = 0.0

    print "X range (", x1, "~", x2, "), Y range(", y1, "~ ", y2, "), Z:", z
    width = img.GetImageXdim()
    height = img.GetImageYdim()
    print "sqrDiff w/h:", width, height
    imData = img.GetPlaneData(0)
    
    for n in range (y1,y2):
        for m in range (x1,x2):
            val = imData[int(n*width + m)]
            if val >0:
                diff = avg-val
                avg2 += diff*diff
                counter += 1

    if counter >0:
        avg2 /= counter    
        avg2 = math.sqrt(avg2)    
    return float(avg2)

#note: we may need to update this to handle different binning modes (currently configured for x11)    
def GetSensorStripAverage(img,sqrdiff=None,average=None):
    global sensorStrips
    global sensorX
    global sensorY
    global numSensors
    global sensorsH
    global sensorsV

    if sqrdiff is None:
        sqrdiff = 16383
    if average is None:
        average = 4000
    
    numPix = 0
    xmin1 = 0
    xmax1 =0
    xmin2 = 0
    xmax2 = 0
    ymin = 0
    ymax = 0
    
    width = img.GetImageXdim()
    height = img.GetImageYdim()
    
    stripLength = int(width/sensorsH/sensorStrips)
    results = [0.0 for i in range(sensorStrips*numSensors)]
    
    counter = 0
    limit = int(average+(3*sqrdiff))
    
    imData = img.GetPlaneData(0)
    nData = numpy.array(imData)
    nData = numpy.reshape(nData,[height,width])
    
    for i in range(sensorsV):
        ymin = int(sensorY*0.2)
        ymax = sensorY - ymin
        if ymin < 0:
            ymin = 0
        if ymax > sensorY:
            ymax = sensorY-1
            
        if(model == 2321):
            ymin = int(height * 0.025)
            ymax = int(height * 0.25)
            
        ymin += i*sensorY
        ymax += i*sensorY
        y = ymin
        
        xmin1 = 5
        xmin2 = int(((sensorX/sensorStrips)/2)+5)
        xwidth = int(((sensorX/sensorStrips)/2)-10)
        
        for n in range(sensorStrips*sensorsH):
            val = 0
            #left of readout channel
            xminL = n*(sensorX/sensorStrips)+xmin1
            xmaxL = xminL+xwidth          
            left = nData[ymin:ymax,xminL:xmaxL]
            #right of readout channel
            xminR = n*(sensorX/sensorStrips) + xmin2
            xmaxR = xminR+xwidth            
            right = nData[ymin:ymax,xminR:xmaxR]
            
            total = numpy.add(left,right)
            results[int(i*len(results)/2 + n)] = numpy.average(total)/2
    return results
    
    
def SendConfig(det,offsets,model):
    if(model != 1313):
        for index, n in enumerate(offsets):
            sensor = int(index/6+1)
            strip = (index%6)+1
            address = 20+strip
            print("Address: %d, value: %d, sensor: %d" %(address,n,sensor))
            det.WriteRegister(address,n,sensor)
            time.sleep(0.2)
            
    else:
        iAddr1313DAC_OFFSET1_G_H1_block = 365
        iDACOffsetAddr = iAddr1313DAC_OFFSET1_G_H1_block
        for index, n in enumerate(offsets):
            iDACOffsetAddr = iAddr1313DAC_OFFSET1_G_H1_block + index * 2;
            det.WriteRegister(iDACOffsetAddr,n,1)
            time.sleep(0.2)
            
            
def SingleImage(det, count):
    print("Capturing Image: " + str(count))
    img = DexelaPy.DexImagePy()
    det.Snap(count, 1000)    
    det.ReadBuffer(count,img);
    img.UnscrambleImage() 

    #filename = ('Image %d.smv' % count)
    #img.WriteImage(filename)
        
    return img

def PrintStripAvg(imgNum, avgs):
    try:
        msg = "img" +str(imgNum) +" strip average: "
        for n in range(len(avgs)):
            #msg +=str(avgs[n])
            msg +="  " + str("%.0f" % avgs[n]) 

        print msg
        return
    except :
        return

def GetDarkOffsets(det,filename,StartDarkCalibration, StopDarkCalibration, darkOffset):
    print "GetDarkOffsets\n"
    
    global model
    
    det.SetBinningMode(DexelaPy.bins.x11)
    regVal = det.ReadRegister(1)
    det.WriteRegister(1,4)
    expTimeOrig = det.GetExposureTime()
    det.SetExposureTime(0)
    model = det.GetModelNumber()
    print "Detector Model:", model;
    
    ConfigureGlobals()
    print "done ConfigGlobals"
    if model != 1313:
    
        if model == 2923 or model == 2321:
            lookup = [0,1,2,3,4,5,17,16,15,14,13,12,23,22,21,20,19,18,6,7,8,9,10,11]
    
        if model == 2315 or model == 2307:
            lookup = [6,7,8,9,10,11,0,1,2,3,4,5]
    
        if model == 1512 or model == 1207:
            lookup = [i for i in range(6)]
   
    offsets = [StartDarkCalibration for i in range(numSensors*sensorStrips) ]
    SendConfig(det,offsets,model)
    print "done setConfig"
    time.sleep(1)
    img = SingleImage(det,1)
    print "done SingleImage"
    res = GetSensorStripAverage(img)
    print "done GetSensorStripAverage"
    sqrd = sqrDiff(img,res[1],0,0,img.GetImageXdim(),img.GetImageYdim(),0)
    print "done sqrDiff"
    PrintStripAvg(1, res)

    res = GetSensorStripAverage(img,sqrd,res[1])
    calibdif = StopDarkCalibration-StartDarkCalibration
    offsets = [StopDarkCalibration for i in range(numSensors*sensorStrips)]
    SendConfig(det,offsets,model)
    time.sleep(1)
    img2 = SingleImage(det,2)
    
    res2 = GetSensorStripAverage(img2,sqrd,res[1])
    slope = [0.0 for i in range(len(offsets))]
    PrintStripAvg(2, res2)

    for n in range(len(offsets)):
        dif = res2[lookup[n]]-res[lookup[n]]
        if dif == 0:
            return
        slope[n] = dif/calibdif
        
        dif2 = darkOffset - res[lookup[n]]
        
        cor = int(dif2/slope[n] + 0.5)
        
        offsets[n] = StartDarkCalibration + cor
        if offsets[n]<0:
            offsets[n] = 0
            
    SendConfig(det,offsets,model)
    time.sleep(1)
    img3 = SingleImage(det,3)
    res3 = GetSensorStripAverage(img3,sqrd,res[1])
    PrintStripAvg(3, res3)
    
    for n in range(len(offsets)):
        offsets[n] += 10
        
    SendConfig(det,offsets,model)
    time.sleep(1)
    img4 = SingleImage(det,4)
    res4 = GetSensorStripAverage(img4,sqrd,res[1])
    PrintStripAvg(4, res4)
    
    for n in range(len(offsets)):
        dif = res4[lookup[n]]-res3[lookup[n]]
        if dif == 0:
            return
        slope[n] = dif/10
        dif2 = darkOffset - res3[lookup[n]]
        cor = int(dif2/slope[n] + 0.5)
        offsets[n] = offsets[n] - 10 + int(cor*0.5)
        print n, "th offsets=", offsets[n]
        if offsets[n]<0:
            offsets[n] = 0
    
    SendConfig(det,offsets,model)
    time.sleep(1)
    SingleImage(det,5)

    print "Saving ConfigFile", offsets,filename,model
    SaveConfigFile(offsets,filename,model)
    print "Save ConfigFile", offsets,filename,model
     
    det.SetExposureTime(expTimeOrig)
    det.WriteRegister(1,regVal)
    print("success!")
    return 

def StripAverage(img, filename):
    results = GetSensorStripAverage(img)
    
    counter = 0    
    #mytext = ["Sensor 1\r\n\r\n","Sensor 2\r\n\r\n","Sensor 3\r\n\r\n","Sensor 4\r\n\r\n"]
    mytext = ["","","",""]
       
    
    for n in range(sensorsH):
        for m in range(sensorsV):
            if n == 0 and m == 0:
                counter = 0
            if n == 0 and m == 1:
                counter = 1
            if n == 1 and m == 0:
                counter = 3
            if n == 1 and m == 1:
                counter = 2
   
            for h in range(sensorStrips):
                offset = n*sensorStrips
                offset += m*sensorStrips*sensorsH
                if(h <sensorStrips-1):
                    val =  "%.2f," % results[h+offset]
                else:
                    val =  "%.2f" % results[h+offset]
                mytext[counter] += val
            mytext[counter] += "\r\n" 
    mytext[-1] += "\r\n\r\n"    
    return mytext
            
def LinMeasurement(det,img,filename):
    global model
    model = det.GetModelNumber()    
    ConfigureGlobals()
    res = StripAverage(img,filename)
    if(os.path.isfile(filename)):
        f = open(filename,'a')
    else:
        f = open(filename,'w+')
    for p in range(len(res)):
        if res[p] != None:
            f.write(str(res[p]))
    f.close()
    return






def GetSensorRegionAverage(img):
    global sensorStrips
    global sensorX
    global sensorY
    global numSensors
    global sensorsH
    global sensorsV
   
    width = img.GetImageXdim()
    height = img.GetImageYdim()
    stripWidth =int(width/(numSensors*sensorStrips))
    stripHeight =int(height/sensorsV)
    
    xstart =5  #skip pixels on the edge
    ystart =10 #skip pixels on the edge

    vsBlocks =3 #divide each strip in to 3-vertical pieces.
    blockWidth =stripWidth -10
    blockHeight=int(stripHeight/vsBlocks)-20

    results = [0.0 for i in range(sensorStrips*numSensors*vsBlocks)]
    #print "num of resutls:", len(results)
    
    counter = 0
    
    imData = img.GetPlaneData(0)
    nData = numpy.array(imData)
    nData = numpy.reshape(nData,[height,width])
    
    
    for i in range(sensorsV):
        for j in range(vsBlocks):
            ymin =ystart + i*stripHeight + j*blockHeight
            ymax =ymin + blockHeight
            
            if ymin < 0:
                ymin = 0
            if ymax > sensorY:
                ymax = sensorY-1
                
            if(model == 2321):
                print "Utils.GetSensorRegionAverage not support model 2321. Stop test!"
                sys.exit();
                
           
            xmin1 = xstart
            xmin2 = xmin1 + int(blockWidth/2)
            xwidth =int(blockWidth/2)
            
            for n in range(sensorStrips*sensorsH):
                val = 0
                ### left of readout channel
                xminL = n*stripWidth+xmin1
                xmaxL = xminL+xwidth          
                left = nData[ymin:ymax,xminL:xmaxL]
                #print "ymin ~ymax, xmimL ~xmaxL", ymin, ymax, xminL, xmaxL
                
                #### right of readout channel
                xminR = n*stripWidth + xmin2
                xmaxR = xminR+xwidth            
                right = nData[ymin:ymax,xminR:xmaxR]
                #print "ymin ~ymax, xmimR ~xmaxR", ymin, ymax, xminR, xmaxR
                
                total = numpy.add(left,right)
                #print "index in result", counter
                results[counter] = numpy.average(total)/2

                counter +=1

    return results
    
    
def GetRegionAverage(det, img):
    global model

    model =det.GetModelNumber()

    ConfigureGlobals()

    res =GetSensorRegionAverage(img)
    PrintStripAvg(1, res)

    return res

def GetImageAverage(det, img):
    global model

    model =det.GetModelNumber()

    ConfigureGlobals()

    res =GetSensorRegionAverage(img)

    cnt =len(res)

    if cnt==0:
        return 0
    
    avg=0
    for ii in range(0,cnt):
        avg +=res[ii]
    
    avg =avg/cnt
    print "Image Avg: %.1f" % avg
    return avg
    
    


def setGlobals(theModel):
    global model

    model =theModel
    print "model ", model
    
    ConfigureGlobals()

    if model != 1313:
    
        if model == 2923 or model == 2321:
            lookup = [0,1,2,3,4,5,17,16,15,14,13,12,23,22,21,20,19,18,6,7,8,9,10,11]
    
        if model == 2315 or model == 2307:
            lookup = [6,7,8,9,10,11,0,1,2,3,4,5]
    
        if model == 1512 or model == 1207:
            lookup = [i for i in range(6)]
    


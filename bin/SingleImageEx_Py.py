import DexelaPy
import msvcrt
       

def AcquireSingleImage(detector):
    expMode = DexelaPy.ExposureModes.Expose_and_read
    binFmt = DexelaPy.bins.x11
    wellMode = DexelaPy.FullWellModes.High
    trigger = DexelaPy.ExposureTriggerSource.Internal_Software
    exposureTime = 250
    img = DexelaPy.DexImagePy()

    print("Connecting to detector...")
    detector.OpenBoard()
    w = detector.GetBufferXdim()
    h = detector.GetBufferYdim()
    print("Initializing detector settings...")
    detector.SetFullWellMode(wellMode)
    detector.SetExposureTime(exposureTime)
    detector.SetBinningMode(binFmt)  
    detector.SetTriggerSource(trigger)
    detector.SetExposureMode(expMode)
    model = detector.GetModelNumber()

    if trigger == DexelaPy.ExposureTriggerSource.Internal_Software:
        print("Press any key to trigger detector!")
        while msvcrt.kbhit() != True:
            pass
        msvcrt.getch()
        
    detector.Snap(1, exposureTime+1000)
    
    print("Grabbed Image!")
    detector.ReadBuffer(1,img);

    img.UnscrambleImage() 

    filename = 'Image_%dx%d.smv' % (img.GetImageXdim(),img.GetImageYdim())
    img.WriteImage(filename)

    print("Image successfully saved!")
    detector.CloseBoard()
    
    return 

try:
    
    print("Scanning to see how many devices are present...")
    scanner = DexelaPy.BusScannerPy()

    count = scanner.EnumerateDevices()
    print("Found %d devices " % count)

    for i in range(0,count):
        info = scanner.GetDevice(i)
        det = DexelaPy.DexelaDetectorPy(info)
        print("Acquiring single image from detector with serial number: %d" % info.serialNum)
        AcquireSingleImage(det)

except DexelaPy.DexelaExceptionPy as ex:
    print("Exception Occurred!")
    print("Description: %s" % ex)
    DexException = ex.DexelaException
    print("Function: %s" % DexException.GetFunctionName())
except Exception:
    print("Exception OCCURRED!")




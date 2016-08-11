import Faxitron
import time

time.sleep(1.0)
kVp = 50.0
xOnTimeSecs = 10.00
print("Single exposure with:")
print("KV: " + str(kVp))  
print("X-Ray Time: " + str(xOnTimeSecs))  
    
cab = Faxitron.Cabinet(3)
if(kVp != 0):
    cab.Configure(kVp,xOnTimeSecs*1000)
    if cab.FireXRay() != True:
        print("Faxitron did not fire.")
    cab.WaitXRayOn()
    cab.WaitXRayOff()
    print("Exposure done!")
cab.close()
time.sleep(0.5)

class DexElement:
    'Dexela XML Element'
    
    def __init__(self, isCommand, command,xOnTimeSecs,numberOfExposures,Binning,t_expms,trigger,kVp,mA,comment,fullWell,darkCorrect,gainCorrect,defectCorrect):
        self.isCommand = isCommand
        self.command = command
        self.xOnTimeSecs = xOnTimeSecs
        self.numberOfExposures = numberOfExposures
        self.Binning = Binning
        self.t_expms = t_expms
        self.trigger = trigger
        self.kVp = kVp
        self.mA  = mA
        self.comment = comment
        self.fullWell = fullWell
        self.darkCorrect = darkCorrect
        self.gainCorrect = gainCorrect
        self.defectCorrect = defectCorrect
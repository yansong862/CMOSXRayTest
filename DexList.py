import xml.etree.ElementTree as ET
import DexElement
import DexelaPy

class DexList:
    'Dexela XML List'
    
    def __init__(self, filename):
        
        self.xmlList = []
        
        tree = ET.parse(filename)
        root = tree.getroot()
        for child in root:
            command = child.find('command').text
            comment = child.find('comment').text
            numberOfExposures = int(child.find('numberOfExposures').text)
            t_expms = int(child.find('t_expms').text)
            kVp = int(child.find('kVp').text)
            mA = int(child.find('mA').text)
            xOnTimeSecs = int(child.find('xOnTimeSecs').text)

            if(child.find('isCommand').text == 'true'):
                isCommand = True
            else:
                isCommand = False
             
            if(child.find('darkCorrect').text == 'true'):
                darkCorrect = True
            else:
                darkCorrect = False
            
            if(child.find('gainCorrect').text == 'true'):
                gainCorrect = True
            else:
                gainCorrect = False

            if(child.find('defectCorrect').text == 'true'):
                defectCorrect = True
            else:
                defectCorrect = False
             
            binTxt = child.find('Binning').text
            if binTxt == 'x11':
                Binning = DexelaPy.bins.x11
            elif binTxt == 'x12':
                Binning = DexelaPy.bins.x12
            elif binTxt == 'x14':
                Binning = DexelaPy.bins.x14
            elif binTxt == 'x21':
                Binning = DexelaPy.bins.x21
            elif binTxt == 'x22':
                Binning = DexelaPy.bins.x22
            elif binTxt == 'x24':
                Binning = DexelaPy.bins.x24
            elif binTxt == 'x41':
                Binning = DexelaPy.bins.x41
            elif binTxt == 'x42':
                Binning = DexelaPy.bins.x42
            elif binTxt == 'x44':
                Binning = DexelaPy.bins.x44
            elif binTxt == 'ix22':
                Binning = DexelaPy.bins.ix22
            elif binTxt == 'ix42':
                Binning = DexelaPy.bins.ix44
            else:
                Binning = DexelaPy.bins.BinningUnknown
                
            
            trigText = child.find('trigger').text
            if trigText == 'Internal_Software':
                trigger = DexelaPy.ExposureTriggerSource.Internal_Software
            elif trigText == 'Ext_neg_edge_trig':
                trigger = DexelaPy.ExposureTriggerSource.Internal_Software
            elif trigText == 'Ext_Duration_Trig':
                trigger = DexelaPy.ExposureTriggerSource.Internal_Software
            
            wellText = child.find('fullWell').text
            if wellText == 'High':
                fullWell = DexelaPy.FullWellModes.High
            else:
                fullWell = DexelaPy.FullWellModes.Low
            
            el = DexElement.DexElement(isCommand, command,xOnTimeSecs,numberOfExposures,Binning,t_expms,trigger,kVp,mA,comment,fullWell,darkCorrect,gainCorrect,defectCorrect)
            
            self.xmlList.append(el)
# -*- coding: utf-8 -*-
"""
Created on Mon Feb  1 15:33:07 2021

@author: Tfahry

Unit diagram reader.

the (new) standardized UD entry format - transform everything into this

UDEntry = 
{location : 
 arrTime  :
 arrHeadcode:
 depTime  :
 depHeadcode :
 activity :     
 }

"""

from pandas import read_excel
from NRFunctions import timeStandardiser
from lxml import etree as et

class Reader:
    def __init__(self, pathToUD):
        self.EmptyFill = 'UDNONE'
        self. udEntryFormat = {
         'location'     :   None,
         'arrTime'      :   None,
         'arrHeadcode'  :   None,
         'depTime'      :   None,
         'depHeadcode'  :   None,
         'activity'     :   None,
         'excelRow'     :   None
         }
        
        self.standardised = True #default value
        self.hasExcelRows = False
        
        self.pathToUD = pathToUD
        self.ud = self.Parse(self.pathToUD)
    
    def Parse(self, pathToUD):
        raise NotImplementedError('Parse() not implemented')
        
        
'''
Avanti Reader: user needs to drag and drop Word UD into Excel, trim off all rows up to "Diagram:" row 
and trim all columns up to station name
'''    
class Avanti(Reader):
    def __init__(self, pathToUD):
        super().__init__(pathToUD)
        self.hasExcelRows = True
        
    def Parse(self, pathToUD):
        udEntries = []
        z = read_excel(pathToUD, dtype = str, header = None)
            
        #create a list that looks like [(index,row_contents),(index,row_contents) ...]
        z = [j for i,j in z.iterrows()] 

        for idx, row in enumerate(z):
            if idx == len(z)-1: #if second to last row, do nothing
                pass
            else:
                if z[idx+1].iloc[0] == z[idx].iloc[0]:
                    
                    udEntry = self.udEntryFormat.copy()
                    
                    udEntry['location']     = z[idx].iloc[0]
                    udEntry['arrTime']      = z[idx].iloc[1]
                    udEntry['arrHeadcode']  = z[idx-1].iloc[3]
                    udEntry['depTime']      = z[idx+1].iloc[2]
                    udEntry['depHeadcode']  = z[idx+1].iloc[3]
                    udEntry['activity']     = z[idx].iloc[4]
                    udEntry['excelRow']     = str(idx+1)
                    
                    udEntries.append(udEntry)
                    #sheet.range(f'A{idx}').color = (140, 184, 255)

        
        for idx, entry in enumerate(udEntries):
            #convert all times into hh:mm:ss format understood by RailSys
            udEntries[idx]['arrTime'] = timeStandardiser(udEntries[idx]['arrTime'])
            udEntries[idx]['depTime'] = timeStandardiser(udEntries[idx]['depTime'])
            
            if udEntries[idx]['activity'] == 'REVRSE':
                udEntries[idx]['activity'] = 'turnaround'
            else:
                udEntries[idx]['activity'] = None
        
        return udEntries

'''
ScotRail Reader 1: user needs to drag and drop Word UD into Excel
'''
class ScotRail(Reader):
    def __init__(self, pathToUD):
        super().__init__(pathToUD)
        self.standardised = False
    def Parse(self, pathToUD):
        return read_excel(pathToUD, usecols=[0,1,2,4], header=0, dtype = str).fillna(self.EmptyFill)


# '''s
# ScotRail Reader 2: user needs to drag and drop Word UD into Excel and trim all rows/cols
# so that Location-Arr-Dep header is in the top left corner
# '''
# class ScotRailDec19(Reader):
#     def __init__(self, pathToUD):
#         super().__init__(pathToUD)
#         self.standardised = False
#     def Parse(self, pathToUD):
#         return read_excel(pathToUD, usecols=[0,1,2,4], header=8, dtype = str).fillna(self.EmptyFill)



# class FTPE(Reader):
#     pass
#     #def Parse(self, pathToUD):
        
        
        #return None

if __name__ == '__main__':
    pass
    #print(ScotRailDec19('udec19.xlsx').ud)
    
    #print(ScotRailECML('u170.xlsx').ud)
    
    #print(Avanti('uAvanti.xlsx').ud)
    
    z = Avanti('uAvanti.xlsx')
    y  = Avanti('uXCdec19.xlsx')


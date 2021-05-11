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


class Reader:
    def __init__(self, pathToUD):
        self.EmptyFill = 'UDNONE'
        self. udEntryFormat = {
         'location'     :   None,
         'arrTime'      :   None,
         'arrHeadcode'  :   None,
         'depTime'      :   None,
         'depHeadcode'  :   None,
         'activity'     :   None}
        
        self.standardised = True #default value
        
        self.ud = self.Parse(pathToUD)
    
    def Parse(self, pathToUD):
        raise NotImplementedError('Parse() not implemented')

'''
ScotRail Reader 1: user needs to drag and drop Word UD into Excel
'''
class ScotRailECML(Reader):
    def __init__(self, pathToUD):
        super().__init__(pathToUD)
        self.standardised = False
    def Parse(self, pathToUD):
        return read_excel(pathToUD, usecols=[0,1,2,4], header=6, dtype = str).fillna(self.EmptyFill)


'''
ScotRail Reader 2: user needs to drag and drop Word UD into Excel
'''
class ScotRailDec19(Reader):
    def __init__(self, pathToUD):
        super().__init__(pathToUD)
        self.standardised = False
    def Parse(self, pathToUD):
        return read_excel(pathToUD, usecols=[0,1,2,4], header=8, dtype = str).fillna(self.EmptyFill)

'''
Avanti Reader 1: user needs to drag and drop Word UD into Excel, trim off all rows up to "Diagram:" row 
and trim all columns up to station name
'''    
class Avanti(Reader):
    def Parse(self, pathToUD):
        udEntries = []
        z = read_excel(pathToUD, dtype = str)
        
        #create a list that looks like [(index,row_contents),(index,row_contents) ...]
        z = [j for i,j in z.iterrows()] 
        
        for idx, row in enumerate(z):
            if idx == len(z)-1: #if second to last row, do nothing
                pass
            else:
                if z[idx+1][0] == z[idx][0]:
                    
                    udEntry = self.udEntryFormat.copy()
                    
                    udEntry['location']     = z[idx][0]
                    udEntry['arrTime']      = z[idx][1]
                    udEntry['arrHeadcode']  = z[idx-1][3]
                    udEntry['depTime']      = z[idx+1][2]
                    udEntry['depHeadcode']  = z[idx+1][3]
                    udEntry['activity']     = z[idx][4]
                    
                    udEntries.append(udEntry)

        #convert all times into hh:mm:ss format understood by RailSys
        for idx, entry in enumerate(udEntries):
            udEntries[idx]['arrTime'] = timeStandardiser(udEntries[idx]['arrTime'])
            udEntries[idx]['depTime'] = timeStandardiser(udEntries[idx]['depTime'])
        
        return udEntries

if __name__ == '__main__':
    pass
    #print(ScotRailDec19('udec19.xlsx').ud)
    
    print(ScotRailECML('u170.xlsx').ud)
    
    print(Avanti('uAvanti.xlsx').ud)
    
#%%


# import pandas as pd


# udEntries = []
# z = read_excel(r'C:\Users\Tfarhy\OneDrive - Network Rail\2020.11.24_XML backend for adding connections\Donovan10 diagrams\Avanti\_uAvanti.xlsx')
# z = [j for i,j in z.iterrows()] #contains a list [(idx,row),(idx,row)...]
# udEntryFormat = {
#          'location'     :   None,
#          'arrTime'      :   None,
#          'arrHeadcode'  :   None,
#          'depTime'      :   None,
#          'depHeadcode'  :   None,
#          'activity'     :   None}

# for idx, row in enumerate(z):
#     if idx == len(z)-1: #if second to last row, do nothing
#         pass
#     else:
#         if z[idx+1][0] == z[idx][0]:
            
#             udEntry = udEntryFormat.copy()
            
#             udEntry['location']     = z[idx][0]
#             udEntry['arrTime']      = z[idx][1]
#             udEntry['arrHeadcode']  = z[idx-1][3]
#             udEntry['depTime']      = z[idx+1][2]
#             udEntry['depHeadcode']  = z[idx+1][3]
#             udEntry['activity']     = z[idx][4]
            
#             udEntries.append(udEntry)

# #convert all times into hh:mm:ss format understood by RailSys
# for idx, entry in enumerate(udEntries):
#     udEntries[idx]['arrTime'] = timeStandardiser(udEntries[idx]['arrTime'])
#     udEntries[idx]['depTime'] = timeStandardiser(udEntries[idx]['depTime'])
    
    
    
    
# #%%
# from NRFunctions import timeHandler
# timeHandler('12:34:578')
   
    
    
    
    
    
    
    
    
    
    
    
    

"""
connectionGenerator Module

Created on Tue Dec  8 15:51:53 2020
Requires Python 3.6 or newer.

@author: tfahry, Network Rail C&CA

This script takes a single structured unit diagram and generates a list of 
connections from it.

This module should not implement any XML logic - any XML operations should 
instead go in RSXParser's function definitions and only be called here.

"""
import xlwings as xw
import RSXParser as rp
from NRFunctions import log, hashfile, ResultType, timeHandler as th

#use datetime library instead
from time import time


#%%
start_time = time() #script timer - debug only

wb = xw.Book('Input.xlsx')
u170 = wb.sheets['u170']

tree = rp.read('longlands.rsx')
className = '170'

removed = 0
for conn in tree.findall('.//entry[@stationID="EDINBUR"]/connection'):
    if className in conn.getparent().attrib['trainTypeId']:
        conn.getparent().remove(conn)
        removed = removed+1

print(f'Removed {removed} class {className} EDINBUR connections.\n')

location = u170.range('A1:A3000').value
arr = u170.range('B1:B3000').value
dep = u170.range('C1:C3000').value
train  = u170.range('E1:E3000').value

location.insert(0,None)
arr.insert(0,None)
dep.insert(0,None)
train.insert(0,None)


u170.range('A1:J3000').color = None

result = ResultType()

stationID = 'EDINBUR'

for i, trainname in enumerate(train):
    if trainname is not None and train[i+1] is not None and train[i+1]!=trainname and location[i+1] in ['Edinburgh',] and location[i]!='Location':
        u170.range(f'A{i+1}:J{i+1}').color = (189, 211, 217)
        
        arrTime = th(arr[i+1])
        depTime = th(dep[i+1])
        
        row = (stationID,train[i],arrTime,train[i+1],depTime)
        
        result.tried.app(row)
        
        try:
            entryArr    = rp.findUniqueEntry(tree,train[i][0:4],stationID,arrTime,-1)
            entryWait   = rp.findUniqueEntry(tree,train[i+1][0:4],stationID,depTime,0)
            
                        
            conn = rp.makecon(entryArr)
            
            if not rp.connectionExists(entryWait, conn):
                entryWait.append(conn)
                entryWait.append(rp.et.Element('plus'))
                result.made.app(row)
            else:
                result.duplicate.app(row)
            
        except ValueError as e:
            result.failed.app(row,f'{e} at row {i+1}')
            
            print (f'{e} at row {i+1}')
             
#%%
print(f'Made {result.made.count} connections out of {result.tried.count} in diagram. Rejected {result.duplicate.count} duplicates and failed {result.failed.count}.')
print(f"--- {time()-start_time} seconds ---")

rp.write(tree,'longlands_progadd.rsx')
print(hashfile('longlands_progadd.rsx'))

resultsheet = wb.sheets['Results']
resultsheet.clear()
y = 1

for attr , colour in zip(['failed','made','duplicate'],[(235, 143, 143),(190, 209, 151),(204, 204, 204)]):
    x = getattr(getattr(result,attr),'get')
    count = getattr(getattr(result,attr),'count')
    
    resultsheet.range(f'A{y}').value = x
    
    if attr == 'failed':
        resultsheet.range(f'F{y}').value = list(zip(getattr(getattr(result,attr),'errors')))
        

    resultsheet.range(f'A{y}:E{y+count-1}').color = colour
    
    y = y+ count

#done

        
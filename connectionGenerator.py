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

from abc import ABC, abstractmethod

import xlwings as xw
import RSXParser as rp
#use datetime library instead
from time import strftime, gmtime ,time
from math import ceil


class Counter:
    def __get__(self,instance,owner):
        return len(instance._contents)

class Getter:
    def __get__(self,instance,owner):
        return instance._contents

class ExitType:
    count = Counter()
    get = Getter()
    
    def __init__(self):
        self._contents = []

    def app(self,value):
        self._contents.append(value)
        
    def __set__(self):
        raise AttributeError('__set__ triggered')

class FailedType(ExitType):
    def __init__(self):
        super().__init__()
        self.errors = []
        
    def app(self,value,errormsg):
        self._contents.append(value)
        self.errors.append(errormsg)

class ResultType:    
    def __init__(self):
        self.tried      = ExitType()
        self.made       = ExitType()
        self.duplicate  = ExitType()
        self.failed     = FailedType()
        
        

def removeNone(cells):
    output = []
    for i in cells:
        if i is not None:
            output.append(i)
    return output

#refactor to use strptime()?
def timeHandler(input):
    if type(input) is str:
        if '&' in input:
            if '½' in input:
                return (input.split('&')[-1]).replace('½',':30')
            else:
                return input.split('&')[-1] + ':00'
            
        elif '½' in input:
            return input.replace('½',':30')    
        else:
            raise ValueError('don\'t know how to handle time string {input}')
        
    elif type(input) is float:
        return strftime('%H:%M:%S',gmtime(ceil(86400*input)))
    else:
        raise TypeError(f'cannot handle time {input} because of type {type(input)}')

#shorthand
th=timeHandler


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

        
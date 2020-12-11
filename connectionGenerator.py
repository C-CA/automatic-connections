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
#use datetime library instead
from time import strftime, gmtime ,time
from math import ceil


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

tree = rp.read('longlands_none.rsx')

location = u170.range('A1:A3000').value
arr = u170.range('B1:B3000').value
dep = u170.range('C1:C3000').value
train  = u170.range('E1:E3000').value

location.insert(0,None)
arr.insert(0,None)
dep.insert(0,None)
train.insert(0,None)


u170.range('A1:J3000').color = None

tried, failed, made = 0,0,0

for i, trainname in enumerate(train):
    if trainname is not None and train[i+1] is not None and train[i+1]!=trainname and location[i+1] in ['Edinburgh',] and location[i]!='Location':
        u170.range(f'A{i+1}:J{i+1}').color = (189, 211, 217)
        arrTime = th(arr[i+1])
        depTime = th(dep[i+1])
        
        
        try:
            entryArr    = rp.findUniqueEntry(tree,train[i][0:4],'EDINBUR',arrTime,-1)
            entryWait   = rp.findUniqueEntry(tree,train[i+1][0:4],'EDINBUR',depTime,0)
            
            conn = rp.makecon(entryArr)
            
            entryWait.append(conn)
            
            made = made+1
            #print(f'made conn between {train[i]} and {train[i+1]} at EDINBUR at {arrTime}')
        except ValueError as e:
            failed = failed + 1
            print (f'{e} at row {i+1}')
            #raise ValueError(f'{e} at row {i+1}')
        
        tried = tried+1
        
print(f'Made {made} connections out of {tried}. Failed {failed}.')

print(f"--- {time()-start_time} seconds ---")

rp.write(tree,'longlands_progadd.rsx')
          
        
        
        
        
        
        
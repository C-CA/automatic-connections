"""
Created on Fri Nov 27 17:06:59 2020

Function container module for common C&CA Python utility functions.

@author: tfahry
"""
#%%
"""
External dependencies.
"""

import hashlib
import humanhash
from math import ceil
from time import strftime, strptime, gmtime 
from datetime import datetime


#%%
"""
Class definitions.
"""

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
        self.failed     = ExitType()

#%%
"""
Function definitions.
"""

def log(func):
    def wrapper(*args, **kwargs):
        result  = func(*args, **kwargs)
        print(f'{func.__name__} called with {args}, {kwargs} and returned {result}')
        return result
    return wrapper


def hashfile(file):
    with open(file, "rb") as f:
        file_hash = hashlib.md5()
        while chunk := f.read(8192):
            file_hash.update(chunk)
    return humanhash.humanize(file_hash.hexdigest(), words = 6)


def removeNone(cells):
    output = []
    for i in cells:
        if i is not None:
            output.append(i)
    return output

def timeStandardiser(input):
    input = str(input)
    for time_format in ['%H:%M:%S','%H.%M','%H+%M','%H:%M','%H']:
        try:
            return datetime.strptime(input,time_format).strftime('%H:%M:%S')
        except ValueError:
            pass
    return(f'could not parse time string {input}')
    #raise ValueError(f'cannot handle format of time string {input}')
    
    

#refactor to use strptime()? also use datetime instead of time
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
            for timeFormat in ['%H:%M:%S','%H:%M']:
                try:
                    return strftime('%H:%M:%S',strptime(input,timeFormat))
                except ValueError:
                    continue    
                raise ValueError(f'don\'t know how to handle time string {input}')
    elif type(input) is float:
        return strftime('%H:%M:%S',gmtime(ceil(86400*input)))
    else:
        raise TypeError(f'cannot handle time {input} because of type {type(input)}')
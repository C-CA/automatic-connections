# -*- coding: utf-8 -*-
"""
Created on Mon Feb  1 15:33:07 2021

@author: Tfahry

Unit diagram reader.
"""
from pandas import read_excel

class Reader:
    def __init__(self,pathToUD):
        self.EmptyFill = 'UDNONE'
        self.ud = self.Parse(pathToUD)
    
    def Parse(self, pathToUD):
        raise NotImplementedError('Parse() not implemented')


class ScotRailReader(Reader):
    def Parse(self, pathToUD):
        return read_excel(pathToUD, usecols=[0,1,2,4], header=6, dtype = str).fillna(self.EmptyFill)

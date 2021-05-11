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

from NRFunctions import hashfile, ResultType, timeHandler as th
#from transitiontime import calculateTurnaround
import RSXParser as rp

#use datetime library instead
from time import time

def GenerateConnections(tree, DiagramObject, stationID, stationName):
    result = ResultType()
    made_so_far = set()
    
    if DiagramObject.standardised:
        for i, udEntry in enumerate(DiagramObject.ud):
            if udEntry['location'] == stationName:
                
                row = (stationID,
                       udEntry['arrHeadcode'],
                       udEntry['arrTime'],
                       udEntry['depHeadcode'],
                       udEntry['depTime'])
                
                result.tried.app({'row'         :row,
                                  'entryArr'    :None,
                                  'entryWait'   :None,
                                  'conn'        :None})
                try:
                    #TODO provision for the [0:4] index to be variable
                    entryArr    = rp.findUniqueEntry(tree,
                                                     udEntry['arrHeadcode'][0:4],
                                                     stationID,
                                                     udEntry['arrTime'],
                                                     -1) #binding reference to tree
                    
                    entryWait   = rp.findUniqueEntry(tree,
                                                     udEntry['depHeadcode'][0:4],
                                                     stationID,
                                                     udEntry['depTime'],
                                                     0)
                    
                    conn = rp.makecon(entryArr)
                    conn_tuple = (tuple(entryWait.attrib.items()), tuple(conn.attrib.items()))
                    
                    if not rp.connectionExists(entryWait, conn) and not conn_tuple in made_so_far:
                        
                        made_so_far.add(conn_tuple)
                        
                        result.made.app({'row'          :row,
                                          'entryArr'    :entryArr,
                                          'entryWait'   :entryWait,
                                          'conn'        :conn})
                    else:
                        result.duplicate.app({'row'     :row,
                                          'entryArr'    :entryArr,
                                          'entryWait'   :entryWait,
                                          'conn'        :conn})
                except ValueError as e:
                    #FIXME the row number is wrong, because of standardised UD
                    result.failed.app({'row'       :row,
                                       'error'     :f'{e} at row {i+1}'})            
                    print (f'{e} at row {i+1}')     
                    
        return result
    
    else:
        
        location = DiagramObject.ud['Location']
        arr = DiagramObject.ud['Arr']
        dep = DiagramObject.ud['Dep']
        train = DiagramObject.ud['Train']
        
        for i, trainname in enumerate(train):
            if trainname != DiagramObject.EmptyFill and train[i+1]  != DiagramObject.EmptyFill and train[i+1]!=trainname and location[i+1] in [f'{stationName}',] and location[i]!='Location':
                #u170.range(f'A{i+1}:J{i+1}').color = (189, 211, 217)
                
                arrTime = th(arr[i+1])
                depTime = th(dep[i+1])
                
                row = (stationID,train[i],arrTime,train[i+1],depTime)
                
                result.tried.app({'row'         :row,
                                  'entryArr'    :None,
                                  'entryWait'   :None,
                                  'conn'        :None})
                
                try:
                    #TODO provision for the [0:4] index to be variable
                    entryArr    = rp.findUniqueEntry(tree,train[i][0:4],stationID,arrTime,-1) #binding reference to tree
                    entryWait   = rp.findUniqueEntry(tree,train[i+1][0:4],stationID,depTime,0)
                    
                    conn = rp.makecon(entryArr)
                    conn_tuple = (tuple(entryWait.attrib.items()), tuple(conn.attrib.items()))
                    
                    if not rp.connectionExists(entryWait, conn) and not conn_tuple in made_so_far:
                        
                        made_so_far.add(conn_tuple)
                        
                        result.made.app({'row'          :row,
                                          'entryArr'    :entryArr,
                                          'entryWait'   :entryWait,
                                          'conn'        :conn})
                    else:
                        result.duplicate.app({'row'     :row,
                                          'entryArr'    :entryArr,
                                          'entryWait'   :entryWait,
                                          'conn'        :conn})
                except ValueError as e:
                    #FIXME the row number is wrong, because of the pandas readexcel offset
                    result.failed.app({'row'       :row,
                                       'error'     :f'{e} at row {i+1}'})            
                    print (f'{e} at row {i+1}')     
                    
        return result


def AddConnections(result):
    for item in result.made.get:
        
        entryWait = item['entryWait']
        conn = item['conn']
        
        if not rp.connectionExists(entryWait, conn):
            entryWait.append(conn)
        
    #DEBUFGLAG
    #entryWait.append(rp.et.Element('plus'))

#%% read files
if __name__ == '__main__':
    import RSXParser as rp
    from UnitDiagramReader import ScotRailECML
    
    start_time = time() #script timer - debug only
    
    modelname = 'longlands' #east-connecticut-bluebird-maryland-oxygen-ohio for 6-word humanized md5
    inputfile = modelname + '.rsx'
    
    tree = rp.read(inputfile)
    print(f'Reading {inputfile} (hash {hashfile(inputfile)})')
    
    #%% remove connections that match className
    
    className = '170'
    removed = 0
    for conn in tree.findall('.//entry[@stationID="EDINBUR"]/connection'):
    #for conn in tree.findall('//connection'):
        if className in conn.getparent().attrib['trainTypeId']:
            conn.getparent().remove(conn)
            removed = removed+1
    
    print(f'Removed {removed} class {className} EDINBUR connections.\n')
    
    #%% standard function calls
    diagram = ScotRailECML('u170.xlsx')
    stationID = 'EDINBUR'
    stationName = 'Edinburgh'
    
    result = GenerateConnections(tree, diagram, stationID, stationName)
    AddConnections(result)
          
    #%% write output rsx and result sheet
    
    print(f'Made {result.made.count} connections out of {result.tried.count} in diagram. Rejected {result.duplicate.count} duplicates and failed {result.failed.count}.')
    
    
    outputfile = modelname + '_progadd.rsx' #longlands with u170 = nineteen-butter-ten-fanta-nevada-papa
    
    rp.write(tree,outputfile)
    print(f'\nWrote {outputfile} (hash {hashfile(outputfile)})')
    
    print(f"--- {time()-start_time} seconds ---")
    
    # wb = xw.Book('Input.xlsx')
    # resultsheet = wb.sheets['Results']
    # resultsheet.clear()
    # y = 1
    
    # for attr , colour in zip(['failed','made','duplicate'],[(235, 143, 143),(190, 209, 151),(204, 204, 204)]):
    #     x = getattr(getattr(result,attr),'get')
    #     count = getattr(getattr(result,attr),'count')
        
    #     resultsheet.range(f'A{y}').value = x
        
    #     if attr == 'failed':
    #         resultsheet.range(f'F{y}').value = list(zip(getattr(getattr(result,attr),'errors')))
            
    
    #     resultsheet.range(f'A{y}:E{y+count-1}').color = colour
        
    #     y = y+ count
    
    # #done
        

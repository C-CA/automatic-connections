"""
RSXParser Module

Created on Mon Nov 23 11:18:42 2020
Requires Python 3.6 or newer.

@author: tfahry, Network Rail C&CA

This is an XML parser for manipulating RailSys RSX files.

This parser works directly on RSX files. It contains functions that let 'higher' 
parsers add, get, change and remove entries without them having to implement 
any XML logic.

"""
#%%
"""
Imports and function definitions.
"""

from lxml import etree as et
from datetime import datetime,timedelta


#better length function that returns the length of any iterator, not just ones with a __len__().  
#warning: will exhaust any generator passed to it! Is also O(n), compared to len() which is O(1) 
def len2(input): 
    return sum(1 for _ in input)

#returns first object of a list if the list contains one and only one item, otherwise raises exception.
def getIfExistsAndUnique(input):
    if len2(input)  ==1:
        return input[0]
    elif len2(input)==0:
        raise ValueError('nonexistent')
    elif len2(input)>0:
        raise ValueError('non-unique')
    else:
        raise ValueError('unknown error')
#shorthand 
gu = getIfExistsAndUnique


def findUniqueEntry(tree,trainName,stationID,time,index,secondsTolerance = 600): #conArr entries are always the last stop (index = -1), and conWait the first (index = 0)
    timeSearchList = tree.findall(f'.//train[@name="{trainName}"]//entry[@stationID="{stationID}"]')
    
    entries = []
    if timeSearchList !=[]:
        for entry in timeSearchList:
            if abs((datetime.strptime(entry.attrib['departure'],'%H:%M:%S') - datetime.strptime(time,'%H:%M:%S')).total_seconds())<secondsTolerance:
                entries.append(entry)

    #positive index checker
    entries2 = []    
    for entry in entries:
        timetableentries = entry.getparent()
        
        if index ==-1:
            if len2(timetableentries)-1 == list(timetableentries).index(entry):
                entries2.append(entry)
                
        elif index>=0:
            if index == list(timetableentries).index(entry):
                entries2.append(entry)
        else:
            raise TypeError(f'index {index} not understood @ {stationID} @ {trainName} @ {time}')
    
    try:
        return getIfExistsAndUnique(entries2)
    except ValueError as errormsg:
        raise ValueError(f'{trainName} @ {stationID} @ {time} @ index {index} is {errormsg}')
        

def connectionExists(entryWait, conn):
    return any([entry.attrib == conn.attrib for entry in entryWait.getchildren()])


def makecon(entryArr,
            transitionTime='300',
            validityTime='86399',
            operation='turnaround',
            waitForArrTrain='true'):
    
    #Create an Element for our connection.
    connectionTemplate = et.Element('connection') 
    
    #Populate connection attributes.
    connectionTemplate.attrib['trainNumber']    = entryArr.getparent().getparent().attrib['number']
    if 'numbervar' in entryArr.getparent().getparent().attrib:
        connectionTemplate.attrib['numbervar'] = entryArr.getparent().getparent().attrib['numbervar']
        
    connectionTemplate.attrib['transitionTime'] = transitionTime
    connectionTemplate.attrib['validityTime']   = validityTime
    connectionTemplate.attrib['operation']      = operation    
    if 'pattern' in entryArr.getparent().getparent().attrib:
        connectionTemplate.attrib['trainPattern']   = entryArr.getparent().getparent().attrib['pattern']
        
    connectionTemplate.attrib['stationId']      = entryArr.attrib['stationID']
    connectionTemplate.attrib['trainDeparture'] = entryArr.attrib['departure']
    connectionTemplate.attrib['waitForArrTrain']= waitForArrTrain
    
    return connectionTemplate


def read(filename): #reads an XML and returns an ElementTree object
    parser = et.XMLParser(remove_blank_text=True)
    return et.parse(f'C:\\Users\\Tfarhy\\OneDrive - Network Rail\\2020.11.24_XML backend for adding connections\\{filename}', parser)
    
    
def write(tree,filename):
    et.indent(tree,space='\t')
    tree.write(file = filename, pretty_print=True, xml_declaration=True, encoding="UTF-8", standalone="yes")
    
    
#%%
"""
Demo code. Will not be run if imported as a module.
"""

if __name__ == '__main__':
    modelname = 'mml'
    inputfilename   = modelname + '.rsx'
    outputfilename  = modelname + '_progadd.rsx'
    
    tree = read(inputfilename)
    conns = tree.findall('.//connection')
    
    print(f'{len2(conns)} connections found in {inputfilename}.')
    
    #exclude =   ['1P11','1P31'] #trains that are already known to be non-unique due to user error (duplication), whose connections will raise ValueError in next cell
    exclude =   []
    
    conns2 =    [] #connections that did not raise ValueError in this cell that will be deleted for recursive testing and whose attributes will be appended to [output]
    output =    []
    faulty0 =   [] #connections whose trains raised ValueError in THIS cell due to non-uniquness based on train *number* (possibly due to duplication)
    
    for conn in conns:
        try:
            if 'numbervar' in conn.attrib:
                pt = gu(tree.findall(f'.//train[@number="{conn.attrib["trainNumber"]}"][@numbervar = "{conn.attrib["numbervar"]}"]//entry[@stationID="{conn.attrib["stationId"]}"][@departure="{conn.attrib["trainDeparture"]}"]'))
            else:
                pt = gu(tree.findall(f'.//train[@number="{conn.attrib["trainNumber"]}"]//entry[@stationID="{conn.attrib["stationId"]}"][@departure="{conn.attrib["trainDeparture"]}"]'))
    
                
            if pt.getparent().getparent().attrib['name'] not in exclude:
                
                conns2.append(conn)
                output.append((conn.attrib['stationId'],
                               pt.getparent().getparent().attrib['name'],
                               conn.attrib['trainDeparture'],
                               conn.getparent().getparent().getparent().attrib['name'],
                               conn.getparent().attrib['departure'],
                               conn.attrib['transitionTime'],
                               conn.attrib['validityTime'],
                               conn.attrib['operation']))
        except ValueError:
            faulty0.append(conn)        
            print('warning: ValueError raised in cell 0')
    
    for conn in conns2:
        conn.getparent().remove(conn)
    
    print(f'{len2(faulty0)} connections were non-unique based on number and were not removed.')
    print(f'{len2(conns2)} connections removed and attributes recorded.\n')
#%%
    progconns = []
    faulty =    [] #connections whose trains raised ValueError from FindUniqueEntry due to non-uniquness based on train *name*
    
    for row in output:
        try:
            progconn = makecon(entryArr = findUniqueEntry(tree,row[1],row[0],row[2],-1),
                               transitionTime = row[5],
                               validityTime = row[6],
                               operation = row[7])
            
            entryWait = findUniqueEntry(tree,row[3],row[0],row[4],0)
            if not connectionExists(entryWait,progconn):
                entryWait.append(progconn)
                progconns.append(progconn)
            else:
                print(f'warning: {row} is duplicate')
                
        except ValueError:
            faulty.append(row)
            
    print(f'{len2(faulty)} connections were non-unique based on name and were not generated.')
    print(f'{len2(progconns)} connections generated and added to tree.')
    
    write(tree, outputfilename)
    
    print(f'Tree written to {outputfilename}.')
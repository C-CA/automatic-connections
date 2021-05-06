# -*- coding: utf-8 -*-
"""
Created on Tue Feb  9 10:21:22 2021

@author: Tfahry

Connection Macro Main module

Top-level module containing bindings of PyQt buttons to Connection macro modules. Run Pyinstaller on this file.
"""
#%%
import sys

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem
from PyQt5 import QtCore

from ConnectionMacroUI import Ui_MainWindow
#pyuic5 '.\qt\ConnectionMacroUI.ui' -o ConnectionMacroUI.py'

from connectionGenerator import GenerateConnections, AddConnections
import UnitDiagramReader
from RSXParser import read, write
from NRFunctions import ResultType, hashfile

from copy import deepcopy

class Window(QMainWindow, Ui_MainWindow):
    def __init__(self, parent = None):
        super().__init__(parent)
        self.setupUi(self)
        self.connectSignalsSlots()
        self.result = ResultType()
        
        self.statusBar.showMessage('Connection Macro pre-alpha v0.2')
        
        self.available_ud_readers = [cls.__name__ for cls in UnitDiagramReader.Reader.__subclasses__()]
        
        for reader in self.available_ud_readers:
            self.udselector.addItem(reader)
        
    def connectSignalsSlots(self):
        pass
    
    def debugbutton_clicked(self):
        pass
    
    def rsxbrowse_clicked(self):
        self.options = QFileDialog.Options()
        self.pathToRSX, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*)", 
                                                        options=self.options)
        
        if self.pathToRSX:
            self.lineEdit.setText(self.pathToRSX)
    def udbrowse_clicked(self):
        self.options = QFileDialog.Options()
        self.pathToUD, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*)", 
                                                       options=self.options)
        
        if self.pathToUD:
            self.lineEdit_2.setText(self.pathToUD)

    #FIXME findUniqueEntry throws '24:02:01' error from dec19 unit diagram            
    #TODO make entries autofill from saved [currently hardcoded defaults]
    #TODO minimum and maximum times
    #TODO make failed conns exportable
    #TODO more meaningful error messages
    #TODO dropdown for Activity column in UI
    def generate_clicked(self):
        self.console.append('Making connections...')
        self._tree = read(self.lineEdit.text())
        self.tree = deepcopy(self._tree)
        self.diagram = getattr(UnitDiagramReader,self.udselector.currentText())(self.lineEdit_2.text())
        self.tiploc = self.tiplocbox.text()
        self.stationname = self.stationnamebox.text()
        self.result = GenerateConnections(tree=self.tree, DiagramObject=self.diagram, stationID=self.tiploc, 
                                          stationName = self.stationname)
        
        self.console.append(f'Made {self.result.made.count} connections out of {self.result.tried.count} in diagram. Rejected {self.result.duplicate.count} duplicates and failed {self.result.failed.count}.')
        
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), f'Made [{self.result.made.count}]')
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), f'Duplicates in {self.lineEdit.text().split("/")[-1]} (will not add) [{self.result.duplicate.count}]')
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), f'Falied [{self.result.failed.count}]')
        
        self.setData(self.tableWidget,self.result.made.get)
        self.setData(self.tableWidget_2,self.result.duplicate.get)
        self.setFailed(self.tableWidget_3,self.result.failed.get)
        
    def setData(self, widget, data):
        
        widget.setRowCount(len(data))
        
        for row, item in enumerate(data):
            attrib = item['conn'].attrib
            
            columnMap = {0:'transitionTime',
                         1:'operation',
                         2:'stationId',
                         4:'trainDeparture'}
                
            for key in columnMap.keys():
                columnMap[key] = attrib[columnMap[key]]
            columnMap[3] = item['row'][1]
            columnMap[5] = item['row'][3]
            columnMap[6] = item['row'][4]
            
            chkBoxItem = QTableWidgetItem()
            chkBoxItem.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            chkBoxItem.setCheckState(QtCore.Qt.Unchecked)       
            widget.setItem(row,7,chkBoxItem)
            
            columnMap[7] = chkBoxItem
            
            for key in columnMap.keys():
                newitem = QTableWidgetItem(columnMap[key])
                widget.setItem(row, key, newitem)
            

            
    def setFailed(self, widget, data):
        widget.setRowCount(len(data))
        columnMap = {}
        for row, item in enumerate(data):
            columnMap[0] = item['error']            
            columnMap[1] = item['row'][0]
            columnMap[2] = item['row'][1]
            columnMap[3] = item['row'][2]
            columnMap[4] = item['row'][3]
            columnMap[5] = item['row'][4]

            for key in columnMap.keys():
                newitem = QTableWidgetItem(columnMap[key])
                widget.setItem(row, key, newitem)
    
        
    def savebutton_clicked(self):
        self.options = QFileDialog.Options()
        self.pathToSave, _ = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","RailSys RSX file (*.rsx)", 
                                                         options=self.options)
        
        if self.pathToSave:
            #make a copy of the tree object and regenerate the Result object since any previous run of the AddConnections function
            #would have modified the elements inside Result 
            self.tree = deepcopy(self._tree)
            self.result = GenerateConnections(tree=self.tree, DiagramObject=self.diagram, stationID=self.tiploc, 
                                          stationName = self.stationname)
            
            #Update self.result object with values from Transition Time and Activity columns in self.tablewidget.
            for row_num in range(self.tableWidget.rowCount()):
                self.result.made._contents[row_num]['conn'].attrib['transitionTime'] = self.tableWidget.item(row_num,0).text()
                self.result.made._contents[row_num]['conn'].attrib['operation'] = self.tableWidget.item(row_num,1).text()            
            
            AddConnections(self.result)
            write(tree = self.tree, filename = self.pathToSave)

            self.console.append(f'\nWrote {self.pathToSave} (hash {hashfile(self.pathToSave)})')
            #black-kilo-triple-social-quebec-quiet

            
def excepthook(type, value, tb):
    win.console.append('******  Error  *********')
    win.console.append(str(type))      
    win.console.append(str(value))      
    win.console.append(str(tb.tb_frame))
    
if __name__ =='__main__':
    sys.excepthook = excepthook
    app = QApplication(sys.argv)
    win = Window()
    win.show()
    sys.exit(app.exec())
    


import xlsxwriter as xls
from script.reading.ReadExcel import *

def getDistList(DataFile):
    rowcount = getLastRowCount(DataFile)
    Distlist = []
    for i in range(2,rowcount):
        Distlist.append(getCellData(DataFile,i,0))
    return Distlist;

def splitExcel(DataFile):
    DistList = getDistList(DataFile)
    for i in range(0, len(DistList)-1):
        wb = xls.Workbook("G:\Mohan\PythonProjects\Covid19Report\data\Results\%s.xlsx" % DistList[i])
        ws = wb.add_worksheet()
        wb.close()
    return;
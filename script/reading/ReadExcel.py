import pandas as pd
import xlrd as xl

def getLastRowCount(DataFileLocation):
    try:
        workbook = xl.open_workbook(DataFileLocation)
        sheet = workbook.sheet_by_index(0)
        TotalRows = sheet.nrows
    except FileNotFoundError:
        TotalRows = 0
        print("Exception Occurs")
    return TotalRows;

def getLastColCount(DataFileLocation):
    try:
        workbook = xl.open_workbook(DataFileLocation)
        sheet = workbook.sheet_by_index(0)
        TotalCols = sheet.ncols
    except FileNotFoundError:
        TotalCols = 0
        print("Exception Occurs")
    return TotalCols;


def getCellData(DataFileLocation,RowNum,ColNum):
    try:
        workbook = xl.open_workbook(DataFileLocation)
        sheet = workbook.sheet_by_index(0)
        CellData = sheet.cell_value(RowNum,ColNum)
    except FileNotFoundError:
        CellData = ""
        print("Exception Occurs")
    return CellData;

def getHeaderRowData(DataFileLocation,RowNum,ColNum):
    try:
        workbook = xl.open_workbook(DataFileLocation)
        sheet = workbook.sheet_by_index(0)
        mapList={}
        for i in range(0,RowNum):
            for j in range(0,ColNum):
                mapList.update({sheet.cell_value(0,j):sheet.cell_value(i,j)})
    except FileNotFoundError:
        print("File Not Found")
    return mapList;


def getRowData(DataFileLocation,RowNum,ColNum):
    try:
        workbook = xl.open_workbook(DataFileLocation)
        sheet = workbook.sheet_by_index(0)
        mapList={}
        for i in range(0,RowNum):
            for j in range(0,ColNum):
                mapList.update({sheet.cell_value(0,j):sheet.cell_value(i+1,j)})
    except FileNotFoundError:
        print("File Not Found")
    return mapList;

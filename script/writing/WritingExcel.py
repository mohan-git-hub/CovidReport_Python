import xlsxwriter as xw
import xlrd as xl
from script.writing.ProcessingFile import *
from script.reading.ReadExcel import *

def writeHeaders():
    DataFile="G:\Mohan\PythonProjects\Covid19Report\data\Covid19Data.xlsx"
    DistList = getDistList(DataFile)
    splitExcel(DataFile)
    LastColCount = getLastColCount(DataFile)
    HeaderList = getHeaderRowData(DataFile,1,8)
    for i in range(0,len(DistList)):
        wb1 = xw.Workbook("G:\Mohan\PythonProjects\Covid19Report\data\Results\%s.xlsx" %DistList[i])
        worksheet = wb1.add_worksheet('Sheet1')
        worksheet.write(0,0,HeaderList.get("District"))
        worksheet.write(0,1,HeaderList.get("Diagnosed cases[a]"))
        worksheet.write(0,2,HeaderList.get("Deaths"))
        worksheet.write(0,3,HeaderList.get("Recovered cases"))
        worksheet.write(0,4,HeaderList.get("Active cases[b]"))
        worksheet.write(0,5,HeaderList.get("Population[1]"))
        worksheet.write(0,6,HeaderList.get("Cases per M"))
        worksheet.write(0,7,HeaderList.get("Last case reported on"))
        wb1.close()
        continue
    return;

def writeDistDetails():
    DataFile="G:\Mohan\PythonProjects\Covid19Report\data\Covid19Data.xlsx"
    workbook = xl.open_workbook(DataFile)
    sheet = workbook.sheet_by_index(0)
    LastRowNum = getLastRowCount(DataFile)
    LastColNum = getLastColCount(DataFile)
    DistList = getDistList(DataFile)

    for i in range(0,len(DistList)):
        Dist = DistList[i]
        for j in range(0,LastRowNum):
            CellData = sheet.cell_value(j, 0)
            if CellData == Dist:
                RowNum = j
                DistDetail = getRowData(DataFile,j,LastColNum)
                wb = xw.Workbook("G:\Mohan\PythonProjects\Covid19Report\data\Results\%s.xlsx" % DistList[i])
                worksheet = wb.add_worksheet('Sheet1')
                worksheet.write(0, 0, "District")
                worksheet.write(0, 1, "Diagnosed cases[a]")
                worksheet.write(0, 2, "Deaths")
                worksheet.write(0, 3, "Recovered cases")
                worksheet.write(0, 4, "Active cases[b]")
                worksheet.write(0, 5, "Population[1]")
                worksheet.write(0, 6, "Cases per M")
                worksheet.write(0, 7, "Last case reported on")
                worksheet.write(1, 0, DistDetail.get("District"))
                worksheet.write(1, 1, DistDetail.get("Diagnosed cases[a]"))
                worksheet.write(1, 2, DistDetail.get("Deaths"))
                worksheet.write(1, 3, DistDetail.get("Recovered cases"))
                worksheet.write(1, 4, DistDetail.get("Active cases[b]"))
                worksheet.write(1, 5, DistDetail.get("Population[1]"))
                worksheet.write(1, 6, DistDetail.get("Cases per M"))
                worksheet.write(1, 7, DistDetail.get("Last case reported on"))
                wb.close()
                continue
    return;


writeDistDetails()

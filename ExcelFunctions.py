#http://msdn.microsoft.com/en-us/library/office/ff194068(v=office.14).aspx
import types
import win32com.client as win32
from win32com.client import constants as c

def openExcelFile(file, visible=False):   #takes in a filename, workSheet name and returns an excel & worksheet object
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = visible
    excel.Workbooks.Open(Filename=file,ReadOnly=0)
    return excel

def makeWorkSheetActive(xl, workSheetName): #nice. Makes worksheet active. Returns a worksheet or false
    try:
        xl.Worksheets(workSheetName).Activate
        return xl.Worksheets(workSheetName)
    except:
        return False

def closeExcelDocument(xl, save=True):
    wb = xl.ActiveWorkbook
    wb.Close(SaveChanges=save)

def lastEmptyRowInColumn(xl, column, checkEntireRowIsBlank=False):
    row = xl.Range(column + str(xl.ActiveSheet.Rows.Count)).End(c.xlUp).Row + 1 #find last row in column with data, go down 1
    if checkEntireRowIsBlank:
        while xl.WorksheetFunction.CountA(xl.Range(str(row) + ":" + str(row))) != 0:  #row has an entry
            row = row + 1   #go down a row
    return row

def lastRowUsedRange(xl):
    return xl.Cells.SpecialCells(c.xlCellTypeLastCell).Row

#WARNING: when sending in a date, change to mm/dd/yy due to "feature" in excel com
def append(file, worksheetname, itemsToAdd, addToColumn="A", checkEntireRowIsBlank=True, goToBottomOfColumn=''): #appends array to the end of an excel file. It checks the row is entirely blank before writing.
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        #excel.Visible = False
        wb = excel.Workbooks.Open(Filename=file,ReadOnly=0)
        ws = excel.Worksheets(worksheetname)
        if goToBottomOfColumn:  #allows you to check the end of one column, but write to another
            rowToWriteIn = lastEmptyRowInColumn(excel, goToBottomOfColumn, checkEntireRowIsBlank)
        else:
            rowToWriteIn = lastEmptyRowInColumn(excel, addToColumn, checkEntireRowIsBlank)
        rangeToWriteTo = addToColumn + str(rowToWriteIn) + ":" + chr(ord(addToColumn) + len(itemsToAdd)-1) + str(rowToWriteIn)  #chr(ord("A")+1)="B"
        ws.Range(rangeToWriteTo).Value2 = [i for i in itemsToAdd]
        wb.Save()
        wb.Close()
        return True
    except:
        return False

#WARNING: when sending in a date, change to mm/dd/yy due to "feature" in excel com
def appendToOpenXl(xl, ws, itemsToAdd, addToColumn="A", checkEntireRowIsBlank=True, goToBottomOfColumn=''): #appends array to the end of an excel file. It checks the row is entirely blank before writing.
    try:
        if goToBottomOfColumn:  #allows you to check the end of one column, but write to another
            rowToWriteIn = lastEmptyRowInColumn(xl, goToBottomOfColumn, checkEntireRowIsBlank)
        else:
            rowToWriteIn = lastEmptyRowInColumn(xl, addToColumn, checkEntireRowIsBlank)
        rangeToWriteTo = addToColumn + str(rowToWriteIn) + ":" + chr(ord(addToColumn) + len(itemsToAdd)-1) + str(rowToWriteIn)  #chr(ord("A")+1)="B"
        ws.Range(rangeToWriteTo).Value2 = [i for i in itemsToAdd]
        return True
    except:
        return False

def autoFill(ws, fromRange, toRange):   #e.g. fromRange="A1:A3"
    #http://msdn.microsoft.com/en-us/library/office/ff195345(v=office.14).aspx
    sourceRange = ws.Range(fromRange)
    fillRange = ws.Range(toRange)
    sourceRange.AutoFill(fillRange)

def autoFillDownFromEnd(xl, ws, column, amount):
    rowEndOfData = lastEmptyRowInColumn(xl, column) - 1
    rangeFrom = column + str(rowEndOfData)
    rangeTo = rangeFrom + ':' + column + str(rowEndOfData+amount)
    autoFill(ws, rangeFrom, rangeTo)

def lastCellValueInColumn(ws, column):
    return ws.Range(column + str(ws.Rows.Count)).End(c.xlUp).Value2 #find last cell value

def lastRowInColumn(ws, column):
    return ws.Range(column + str(ws.Rows.Count)).End(c.xlUp).Row #find last row number
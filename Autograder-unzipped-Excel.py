
# coding: utf-8

# # The Excel Autograder

# In[49]:


# get_ipython().run_cell_magic('javascript', '', "\nJupyter.keyboard_manager.command_shortcuts.add_shortcut('/', {\n    help : 'run all cells',\n    help_index : 'zz',\n    handler : function (event) {\n        IPython.notebook.execute_all_cells();\n        return false;\n    }}\n);")


# ### Importing Libraries

# In[50]:


import numpy as np
from xml.dom import minidom
from prettytable import PrettyTable # pip install https://pypi.python.org/packages/source/P/PrettyTable/prettytable-0.7.2.tar.bz2
import zipfile
import os, shutil
from collections import defaultdict
import re
import math
# import time
import csv
#import enchant # unfortunately only is built for 32-bit python
from difflib import SequenceMatcher
import pandas as pd
# from Exams_Fall2017 import *
import json


# ### Styles Dictionaries

# In[51]:


def borderStyle(position):
    tDic = {}
    tDic["style"] = position.getAttribute("style")
    if position.getElementsByTagName("color")[0]:
        if position.getElementsByTagName("color")[0].hasAttribute("theme"):
            tDic["colorTheme"] = position.getElementsByTagName("color")[0].getAttribute("theme")
        if position.getElementsByTagName("color")[0].hasAttribute("tint"):
            tDic["colorTint"] = position.getElementsByTagName("color")[0].getAttribute("tint")
        if position.getElementsByTagName("color")[0].hasAttribute("rgb"):
            tDic["colorRGB"] = position.getElementsByTagName("color")[0].getAttribute("rgb")
        if position.getElementsByTagName("color")[0].hasAttribute("auto"):
            tDic["colorAuto"] = position.getElementsByTagName("color")[0].getAttribute("auto")
        if position.getElementsByTagName("color")[0].hasAttribute("indexed"):
            tDic["colorI"] = position.getElementsByTagName("color")[0].getAttribute("indexed")
    return tDic


def styleDictionary(styleXML):
    stDic = {}
    styleSheet = styleXML.getElementsByTagName("styleSheet")[0]

    # reading each part and save it to dictionaries
    # Number Formats
    numFmtDic = {}
    numFmtDic = defaultdict(lambda: "System_Defined_Format_ID", numFmtDic) # for the keys that don't exist, they are defined in dxfs
    if styleSheet.getElementsByTagName("numFmts"):
        numFmts = styleSheet.getElementsByTagName("numFmts")[0]        
        for numFmt in numFmts.getElementsByTagName("numFmt"):
            if numFmt.hasAttribute("numFmtId") and numFmt.hasAttribute("formatCode"):
                numFmtDic[numFmt.getAttribute("numFmtId")] = numFmt.getAttribute("formatCode")            
    
    # Fonts
    if styleSheet.getElementsByTagName("fonts"):
        fonts = styleSheet.getElementsByTagName("fonts")[0]
        fontDic = {}
        fId = 0
        for font in fonts.getElementsByTagName("font"):
            tempDic = {}
            if font.getElementsByTagName("b"):
                tempDic["bold"] = True
            if font.getElementsByTagName("u"):
                tempDic["underline"] = True
            if font.getElementsByTagName("i"):
                tempDic["italic"] = True
            if font.getElementsByTagName("strike"):
                tempDic["strike"] = True
            if font.getElementsByTagName("sz"):
                tempDic["size"] = font.getElementsByTagName("sz")[0].getAttribute("val")
            if font.getElementsByTagName("color"):
                tempDic["colorT"] = font.getElementsByTagName("color")[0].getAttribute("theme")
                tempDic["colorRGB"] = font.getElementsByTagName("color")[0].getAttribute("rgb")
            if font.getElementsByTagName("name"):
                tempDic["name"] = font.getElementsByTagName("name")[0].getAttribute("val")
            if font.getElementsByTagName("family"):
                tempDic["family"] = font.getElementsByTagName("family")[0].getAttribute("val")
            if font.getElementsByTagName("scheme"):
                tempDic["scheme"] = font.getElementsByTagName("scheme")[0].getAttribute("val")            
            fontDic[fId] = tempDic
            fId += 1

    # Fills
    if styleSheet.getElementsByTagName("fills"):
        fills = styleSheet.getElementsByTagName("fills")[0]
        fillDic = {}
        fillId = 0
        for fill in fills.getElementsByTagName("fill"):
            tempDic = {}
            pFill = fill.getElementsByTagName("patternFill")[0]
            tempDic["pattern"] = pFill.getAttribute("patternType")        
            # In the manual it said fgColor for Solid patternType and bgColor otherwise should be added, 
            # but for "darkUp" pattern there was no bgColor
            if pFill.getElementsByTagName("fgColor"):
                if pFill.getElementsByTagName("fgColor")[0].hasAttribute("rgb"):
                    tempDic["fgColorRGB"] = pFill.getElementsByTagName("fgColor")[0].getAttribute("rgb")
                if pFill.getElementsByTagName("fgColor")[0].hasAttribute("theme"):
                    tempDic["fgColorTheme"] = pFill.getElementsByTagName("fgColor")[0].getAttribute("theme")
                if pFill.getElementsByTagName("fgColor")[0].hasAttribute("tint"):
                    tempDic["fgColorTint"] = pFill.getElementsByTagName("fgColor")[0].getAttribute("tint")
            if pFill.getElementsByTagName("bgColor"):
                tempDic["bgColorI"] = pFill.getElementsByTagName("bgColor")[0].getAttribute("indexed")      
            fillDic[fillId] = tempDic
            fillId += 1        

    # Borders
    if styleSheet.getElementsByTagName("borders"):
        borders = styleSheet.getElementsByTagName("borders")[0]
        borderDic = {}
        borderId = 0
        for border in borders.getElementsByTagName("border"):
            tempDic = {}
            if border.getElementsByTagName("left"):
                if border.getElementsByTagName("left")[0].hasAttribute("style"):
                    position = border.getElementsByTagName("left")[0]
                    tempDic["left"] = borderStyle(position)
            if border.getElementsByTagName("right"):
                if border.getElementsByTagName("right")[0].hasAttribute("style"):
                    position = border.getElementsByTagName("right")[0]
                    tempDic["right"] = borderStyle(position)
            if border.getElementsByTagName("top"):
                if border.getElementsByTagName("top")[0].hasAttribute("style"):
                    position = border.getElementsByTagName("top")[0]
                    tempDic["top"] = borderStyle(position)
            if border.getElementsByTagName("bottom"):
                if border.getElementsByTagName("bottom")[0].hasAttribute("style"):
                    position = border.getElementsByTagName("bottom")[0]
                    tempDic["bottom"] = borderStyle(position)
            if border.getElementsByTagName("diagonal"):
                if border.getElementsByTagName("diagonal")[0].hasAttribute("style"):
                    position = border.getElementsByTagName("diagonal")[0]
                    tempDic["diagonal"] = borderStyle(position)
            borderDic[borderId] = tempDic
            borderId += 1   
            
            
    # Cell defined Styles   
    if styleSheet.getElementsByTagName("cellStyleXfs"): # By default apply is 1
        cellStyleXfs = styleSheet.getElementsByTagName("cellStyleXfs")[0]
        cellStyleDic = {} 
        xfId = 0
        for cellStyle in cellStyleXfs.getElementsByTagName("xf"):
            tempDic = {}
            if xfId != 0:
                if not cellStyle.getAttribute("applyFont"):
                    fontId = cellStyle.getAttribute("fontId")
                    tempDic["font"] = fontDic[int(fontId)]
                if not cellStyle.getAttribute("applyFill"):
                    fillId = cellStyle.getAttribute("fillId")
                    tempDic["fill"] = fillDic[int(fillId)]
                if not cellStyle.getAttribute("applyBorder"):
                    borderId = cellStyle.getAttribute("borderId")
                    tempDic["border"] = borderDic[int(borderId)]
                if not cellStyle.getAttribute("applyNumberFormat"):
                    numFmtId = cellStyle.getAttribute("numFmtId")
                    if numFmtDic[numFmtId] == "System_Defined_Format_ID":
                        tempDic["numFmt"] = numFmtId
                    else:                
                        tempDic["numFmt"] = numFmtDic[numFmtId] # its ID is a string
                if not cellStyle.getAttribute("applyAlignment"):
                    print("Check for the alignment in cellStyle") # alignment will not be defined here
                if not cellStyle.getAttribute("applyProtection"):
                    print("Check for the protection in cellStyle") # protection will not be defined here
                cellStyleDic[xfId] = tempDic
            xfId += 1


    # Cell Xfs
    if styleSheet.getElementsByTagName("cellXfs"): # By default apply is 0
        cellXfs = styleSheet.getElementsByTagName("cellXfs")[0]
        styleId = 0
        for xf in cellXfs.getElementsByTagName("xf"):
            dic = {}
            fontId = 0
            if not xf.getAttribute("xfId") == '0':
                dic = cellStyleDic[int(xf.getAttribute("xfId"))].copy() # if we don't copy it will change the reference

            if xf.hasAttribute("applyFont"):
                fontId = xf.getAttribute("fontId")
            dic["font"] = fontDic[int(fontId)]

            if xf.hasAttribute("applyNumberFormat"):
                numFmtId = xf.getAttribute("numFmtId")
                if numFmtDic[numFmtId] == "System_Defined_Format_ID":
                    dic["numFmt"] = numFmtId
                else:
                    dic["numFmt"] = numFmtDic[numFmtId]

            if xf.hasAttribute("applyAlignment"):
                if xf.getElementsByTagName("alignment"):
                    tempDic = {}
                    align = xf.getElementsByTagName("alignment")[0]
                    if align.hasAttribute("horizontal"):
                        tempDic["alignH"] = align.getAttribute("horizontal")
                    if align.hasAttribute("vertical"):
                        tempDic["alignV"] = align.getAttribute("vertical") 
                    if align.hasAttribute("indent"):
                        tempDic["alignI"] = align.getAttribute("indent") 
                    dic["align"] = tempDic
                    # might be some other alignment
                else: 
                    dic["align"] = "default" # vertical bottom
                    
            if xf.hasAttribute("applyProtection"):
                if xf.getElementsByTagName("protection"):  
                    protect = xf.getElementsByTagName("protection")[0]
                    if protect.hasAttribute("locked"):
                        if protect.getAttribute("locked") == '0':
                            dic["protect"] = 0
                    else:
                        print("Check for other attributes in the protection in cellXfs")
                else: 
                    dic["protect"] = 1

            if xf.hasAttribute("applyBorder"):
                borderId = xf.getAttribute("borderId")
                dic["border"] = borderDic[int(borderId)]

            if xf.hasAttribute("applyFill"):
                fillId = xf.getAttribute("fillId")
                dic["fill"] = fillDic[int(fillId)]

            stDic[styleId] = dic
            styleId += 1  
    return stDic

def dxfsDictionary(styleXML):
    dxfDic = {}
    styleSheet = styleXML.getElementsByTagName("styleSheet")[0]
    
    if styleSheet.getElementsByTagName("dxfs"):
        dxfs = styleSheet.getElementsByTagName("dxfs")[0]
        dxfId = 0
        for dxf in dxfs.getElementsByTagName("dxf"):
            # dxfs are defined with fonts, fills and borders. Unless, they are changed manually
            dic = {}
            if dxf.getElementsByTagName("font"):
                font = dxf.getElementsByTagName("font")[0]
                tempDic = {}
                if font.getElementsByTagName("b"):
                    tempDic["bold"] = True
                if font.getElementsByTagName("u"):
                    tempDic["underline"] = True
                if font.getElementsByTagName("i"):
                    tempDic["italic"] = True
                if font.getElementsByTagName("strike"):
                    tempDic["strike"] = True
                if font.getElementsByTagName("sz"):
                    tempDic["size"] = font.getElementsByTagName("sz")[0].getAttribute("val")
                if font.getElementsByTagName("color"):
                    tempDic["colorRGB"] = font.getElementsByTagName("color")[0].getAttribute("rgb")
                if font.getElementsByTagName("name"):
                    tempDic["name"] = font.getElementsByTagName("name")[0].getAttribute("val")
                if font.getElementsByTagName("family"):
                    tempDic["family"] = font.getElementsByTagName("family")[0].getAttribute("val")
                if font.getElementsByTagName("scheme"):
                    tempDic["scheme"] = font.getElementsByTagName("scheme")[0].getAttribute("val")            
                dic["font"] = tempDic
                
            if dxf.getElementsByTagName("fill"):
                fill = dxf.getElementsByTagName("fill")[0]   
                if fill.getElementsByTagName("patternFill"):
                    pFill = fill.getElementsByTagName("patternFill")[0]
                    tempDic = {}
                    if pFill.getElementsByTagName("fgColor"):
                        tempDic["fgColorRGB"] = pFill.getElementsByTagName("fgColor")[0].getAttribute("rgb")
                    if pFill.getElementsByTagName("bgColor"):
                        tempDic["bgColorRGB"] = pFill.getElementsByTagName("bgColor")[0].getAttribute("rgb")      
                    dic["fill"] = tempDic
                
            if dxf.getElementsByTagName("border"):
                border = dxf.getElementsByTagName("border")[0]
                tempDic = {}
                if border.getElementsByTagName("left"):
                    if border.getElementsByTagName("left")[0].hasAttribute("style"):
                        position = border.getElementsByTagName("left")[0]
                        tempDic["left"] = borderStyle(position)
                if border.getElementsByTagName("right"):
                    if border.getElementsByTagName("right")[0].hasAttribute("style"):
                        position = border.getElementsByTagName("right")[0]
                        tempDic["right"] = borderStyle(position)
                if border.getElementsByTagName("top"):
                    if border.getElementsByTagName("top")[0].hasAttribute("style"):
                        position = border.getElementsByTagName("top")[0]
                        tempDic["top"] = borderStyle(position)
                if border.getElementsByTagName("bottom"):
                    if border.getElementsByTagName("bottom")[0].hasAttribute("style"):
                        position = border.getElementsByTagName("bottom")[0]
                        tempDic["bottom"] = borderStyle(position)
                if border.getElementsByTagName("vertical"):
                    if border.getElementsByTagName("vertical")[0].hasAttribute("style"):
                        position = border.getElementsByTagName("vertical")[0]
                        tempDic["vertical"] = borderStyle(position)
                if border.getElementsByTagName("horizontal"):
                    if border.getElementsByTagName("horizontal")[0].hasAttribute("style"):
                        position = border.getElementsByTagName("horizontal")[0]
                        tempDic["horizontal"] = borderStyle(position)
                dic["border"] = tempDic
            
            dxfDic[dxfId] = dic
            dxfId += 1
    return dxfDic


# ### Functions to Generate Shared Formulas

# In[52]:


def findCellParts(string):
    regex = re.compile("[$]")
    string = regex.sub('', string)
    cellParts = re.findall(r"((\w)+(\d)+)+", string)
    wholeCells = []
    for l in range(len(cellParts)):
        wholeCells.append(cellParts[l][0])
    return wholeCells

def findCell(string):
    cs = []
    wholeCells = findCellParts(string)
    for cell in wholeCells:
        c = ''
        d = ''
        cAlfa = 0
        for letter in cell:
            if letter.isalpha():
                c += letter
                cAlfa += 1
            elif letter.isdigit():
                d += letter
            else:
                cs.append(c)
                cs.append(int(d))
                cAlfa = 0
                c = ''
                d = ''
        cs.append(c)
        cs.append(int(d))
    return cs # column, row, column, row, ..

def changeFormulaCol(formula,number):
    wholeCells = findCellParts(formula)
    for cell in wholeCells:
        c = findCell(cell)
        c[1] = c[1] + number
        newCell = c[0]+str(c[1])
        formula = re.sub(cell, newCell, formula)
    return formula

def changeFormulaRow(formula,number):
    wholeCells = findCellParts(formula)
    for cell in wholeCells:
        c = findCell(cell)
        c[0] = chr(ord(c[0].upper()) + number)
        newCell = c[0]+str(c[1])
        formula = re.sub(cell, newCell, formula)
    return formula

def generateFormula(rangeF,cellRef,formula,formulaDic):
    # on a column
    theRange = findCell(rangeF)
    if len(theRange) == 4:
        if theRange[0] == theRange[2]: 
            for i in range(abs(theRange[3]-theRange[1])):
                formulaDic[theRange[0]+str(i+theRange[1]+1)] = changeFormulaCol(formula,i+1)
        # on a row
        if theRange[1] == theRange[3]: 
            for j in range(abs(ord(theRange[0].upper())-ord(theRange[2].upper()))):
                formulaDic[chr(ord(theRange[0].upper())+j+1)+str(theRange[1])] = changeFormulaRow(formula,j+1)
    return formulaDic 


# ### Shared String File

# In[53]:


def sharedString(sharedstr):
    strArr = []
    strings = sharedstr.getElementsByTagName("sst")[0]
    for s in strings.getElementsByTagName("si"):
        if s.getElementsByTagName("t")[0].firstChild:
            strArr.append(s.getElementsByTagName("t")[0].firstChild.data)
        else:
            strArr.append("EMPTY")
    return strArr


# ### Theme

# In[54]:


def getTheme(themeFile):
    theTheme = themeFile.getElementsByTagName("a:theme")[0]
    if theTheme.hasAttribute("name"):
        return theTheme.getAttribute("name")


# ### Workbooks

# In[55]:


def accessWorkbook(workbook):
    if workbook:
        return workbook.getElementsByTagName("workbook")[0]
    
def sheetNames(workbook):
    theWBook = accessWorkbook(workbook)
    sheets = theWBook.getElementsByTagName("sheets")[0].getElementsByTagName("sheet")
    shNames = []
    for sheet in sheets:
        shNames.append(sheet.getAttribute("name"))
    return shNames

def definedNames(workbook):
    theWBook = accessWorkbook(workbook)
    if theWBook.getElementsByTagName("definedNames"):
        res = {} # name : range
        definedNames = theWBook.getElementsByTagName("definedNames")[0]
        for defName in definedNames.getElementsByTagName("definedName"):
            if defName.hasAttribute("name"):
                name = defName.getAttribute("name")
                if defName.firstChild:
                    res[name] = defName.firstChild.data
        return res          


# ### Tables 

# In[56]:


def accessTable(tableFile):
    if tableFile:
        return tableFile.getElementsByTagName("table")[0]

def tableReference(tableFile):
    table = accessTable(tableFile)
    if table:
        if table.getElementsByTagName("autoFilter"):
            if table.getElementsByTagName("autoFilter")[0].hasAttribute("ref"):
                return table.getElementsByTagName("autoFilter")[0].getAttribute("ref")
            
def tableFilters(tableFile):
    table = accessTable(tableFile)
    if table:
        filters = {}
        if table.getElementsByTagName("autoFilter"):
            if table.getElementsByTagName("autoFilter")[0].getElementsByTagName("filterColumn"):
                cols = table.getElementsByTagName("autoFilter")[0].getElementsByTagName("filterColumn")
                for col in cols:
                    colId = col.getAttribute("colId")
                    if col.getElementsByTagName("filters"):
                        fils = []
                        for fil in col.getElementsByTagName("filters")[0].getElementsByTagName("filter"):
                            fils.append(fil.getAttribute("val"))
                    filters[colId] = fils
            return filters                    


# ### WorkSheets

# In[57]:


def accessWorksheet(worksheet):
    if worksheet:
        return worksheet.getElementsByTagName("worksheet")[0]
    
def getWorksheet(worksheets,sheet):
    for i,wsheet in enumerate(worksheets):
        if i == int(sheet[-1])-1:
            return wsheet


# In[58]:


def worksheetMats(worksheets,strArr):
    formulaDic = {}
    sheetMat = {}
    for i,worksheet in enumerate(worksheets):
        flag = 1 #to avoid numpy silent conversion to string, we put None in formula section of cell A1
        sheet = []
        wsheet = accessWorksheet(worksheet)
        sheetData = wsheet.getElementsByTagName("sheetData")[0]
        rows = sheetData.getElementsByTagName("row")
        for row in rows:
            cells = row.getElementsByTagName("c")
            rowList = []
            for cell in cells:
                if cell.hasAttribute("r"):
                    ref = cell.getAttribute("r")
                else:
                    print("No cell reference available")
                if cell.hasAttribute("s"):
                    style = int(cell.getAttribute("s"))
                else:
                    style = 0 # So we don't have None in the cell styles
                if cell.hasAttribute("t"):
                    if cell.getAttribute("t") == "s":
                        if cell.getElementsByTagName("v")[0].firstChild:
                            value = strArr[int(cell.getElementsByTagName("v")[0].firstChild.data)] 
                    elif cell.getAttribute("t") == "str" or cell.getAttribute("t") == "b" or cell.getAttribute("t") == "e":
                        if cell.getElementsByTagName("v")[0].firstChild:
                            value = cell.getElementsByTagName("v")[0].firstChild.data
                    elif cell.getAttribute("t") == "d": # date
                        if cell.getElementsByTagName("v")[0].firstChild:
                            value = cell.getElementsByTagName("v")[0].firstChild.data.split('.')[0]
                    else:
                        print("There is a new type; rather than v, s, str, e and b:",cell.getAttribute("t"),cell.getAttribute("r"))
                elif cell.getElementsByTagName("v"):
                    value = cell.getElementsByTagName("v")[0].firstChild.data
                else:
                    value = None
                if cell.getElementsByTagName("f"):
                    if cell.getElementsByTagName("f")[0].getAttribute("t") == "shared":
                        if cell.getElementsByTagName("f")[0].hasAttribute("ref"):
                            formula = cell.getElementsByTagName("f")[0].firstChild.data
                            rangeF = cell.getElementsByTagName("f")[0].getAttribute("ref")
                            cellRef = cell.getAttribute("r")                            
                            formulaDic = generateFormula(rangeF,cellRef,formula,formulaDic)
                        else:
                            formula = formulaDic[cell.getAttribute("r")]
                    else:
                        formula = cell.getElementsByTagName("f")[0].firstChild.data                    
                else:
                    formula = None
                if flag == 1:
                    flag = 0
                    rowList.append([ref,style,value,None])
                else:
                    rowList.append([ref,style,value,formula])
            if rowList != []:
                sheet.append(rowList)
        sheetMat["sheet{0}".format(i+1)] = np.array(sheet) 
    return sheetMat


def worksheetDims(worksheets):
    dimension = []
    for worksheet in worksheets:
        wsheet = accessWorksheet(worksheet)
        dimension.append(wsheet.getElementsByTagName("dimension")[0].getAttribute("ref"))
    return dimension

def worksheetColsW(worksheets):
    colsWidth = {}
    for i, worksheet in enumerate(worksheets):
        colsW = {}
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("cols"):
            cols = wsheet.getElementsByTagName("cols")[0].getElementsByTagName("col")
            for col in cols:
                colTag = int(col.getAttribute("min"))-1 # min is the first column affected by this 'column info' record.
                if col.hasAttribute("bestFit"):                
                    colsW[colTag] = "bestFit"
                else:
                    colsW[colTag] = col.getAttribute("width")
                if col.getAttribute("min") != col.getAttribute("max"):
                    for j in range(int(col.getAttribute("max"))-int(col.getAttribute("min"))):
                        if col.hasAttribute("bestFit"):                
                            colsW[colTag+j+1] = "bestFit"
                        else:
                            colsW[colTag+j+1] = col.getAttribute("width")
        colsWidth["sheet{0}".format(i+1)] = colsW
    return colsWidth

def worksheetColsWRef(worksheets):
    colsWidth = {}
    for i,worksheet in enumerate(worksheets):
        colsW = {}
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("cols"):
            cols = wsheet.getElementsByTagName("cols")[0].getElementsByTagName("col")
            
            for col in cols:
                colTag = int(col.getAttribute("min"))-1 # min is the first column affected by this 'column info' record.
                colsW[colTag] = col.getAttribute("width")
                if col.getAttribute("min") != col.getAttribute("max"):
                    for j in range(int(col.getAttribute("max"))-int(col.getAttribute("min"))):
                        colsW[colTag+j+1] = col.getAttribute("width")
        colsWidth["sheet{0}".format(i+1)] = colsW
    return colsWidth

def worksheetColHidden(worksheets):
    hidden = {}
    for i,worksheet in enumerate(worksheets):
        hid = []
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("cols"):
            cols = wsheet.getElementsByTagName("cols")[0].getElementsByTagName("col")
            for col in cols:
                colTag = int(col.getAttribute("min"))-1 # min is the first column affected by this 'column info' record. starting from zerp
                if col.hasAttribute("hidden"):
                    hid.append(colTag)
                if col.getAttribute("min") != col.getAttribute("max"):
                    for j in range(int(col.getAttribute("max"))-int(col.getAttribute("min"))):
                        if col.hasAttribute("hidden"):
                            hid.append(colTag+j+1)
        hidden["sheet{0}".format(i+1)] = hid
    return hidden

def worksheetFreeze(worksheets):
    paneDic = {}
    for i,worksheet in enumerate(worksheets):
        dic = {}
        wsheet = accessWorksheet(worksheet)
        sheetView = wsheet.getElementsByTagName("sheetViews")[0].getElementsByTagName("sheetView")[0]
        if sheetView.getElementsByTagName("pane"):
            pane = sheetView.getElementsByTagName("pane")[0]
            if pane.getAttribute("state") == "frozen":
                if pane.hasAttribute("xSplit"):
                    dic["xSplit"] = pane.getAttribute("xSplit")
                if pane.hasAttribute("ySplit"):
                    dic["ySplit"] = pane.getAttribute("ySplit")
            else:
                print("Pane has been used for anything other than showing a frozen area")
                
        paneDic["sheet{0}".format(i+1)] = dic
    return paneDic

def worksheetConditionalFormatting(worksheets):
    # this function ignores if more than one conditional formatting has applied to one column
    # however, in that case we should consider the priorities. Also, it is assumed it didn't apply partially to a column
    cfDic = {}
    for i,worksheet in enumerate(worksheets):
        dic = {}
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("conditionalFormatting"):
            cfs = wsheet.getElementsByTagName("conditionalFormatting")
            for cf in cfs:
                tempDic = {}
                if cf.hasAttribute("sqref"):
                    tempDic["ref"] = cf.getAttribute("sqref")
                    ref = findCell(cf.getAttribute("sqref"))[0]
                cfRule = cf.getElementsByTagName("cfRule")[0] # this can be generalized
                dxfId = cfRule.getAttribute("dxfId")
                tempDic["dxfId"] = dxfId
                if cfRule.hasAttribute("type"):
                    tempDic["type"] = cfRule.getAttribute("type")
                if cfRule.hasAttribute("bottom"):
                    tempDic["bottom"] = cfRule.getAttribute("bottom")
                if cfRule.hasAttribute("percent"):
                    tempDic["percent"] = cfRule.getAttribute("percent")
                if cfRule.hasAttribute("operator"):
                    operator = cfRule.getAttribute("operator")
                    tempDic["operator"] = operator
                    if operator == "between":
                        formula1 = cfRule.getElementsByTagName("formula")[0].firstChild.data
                        formula2 = cfRule.getElementsByTagName("formula")[1].firstChild.data
                        tempDic["formula"] = (formula1,formula2)
                    else:
                        formula = cfRule.getElementsByTagName("formula")[0].firstChild.data
                        tempDic["formula"] = formula
                dic[ref] = tempDic
        if wsheet.getElementsByTagName("extLst"):
            ext = wsheet.getElementsByTagName("extLst")[0].getElementsByTagName("ext")[0]
            if ext.getElementsByTagName("x14:conditionalFormattings"):
                cfs = ext.getElementsByTagName("x14:conditionalFormattings")[0]
                for cf in cfs.getElementsByTagName("x14:conditionalFormatting"):
                    tempDic = {}
                    cfRule = cf.getElementsByTagName("x14:cfRule")[0] 
                    ref = cf.getElementsByTagName("xm:sqref")[0].firstChild.data
                    if cfRule.hasAttribute("type"):
                        tempDic["type"] = cfRule.getAttribute("type")
                    if cfRule.hasAttribute("bottom"):
                        tempDic["bottom"] = cfRule.getAttribute("bottom")
                    if cfRule.hasAttribute("percent"):
                        tempDic["percent"] = cfRule.getAttribute("percent")
                    if cfRule.hasAttribute("operator"):
                        operator = cfRule.getAttribute("operator")
                        tempDic["operator"] = operator
                    if cfRule.getElementsByTagName("xm:f"):
                        formula = cfRule.getElementsByTagName("xm:f")[0].firstChild.data
                        tempDic["formula"] = formula
                    dxfDic = {}
                    if cfRule.getElementsByTagName("x14:dxf"):
                        dxf = cfRule.getElementsByTagName("x14:dxf")[0]
                        if dxf.getElementsByTagName("font"):
                            dxfDic["colorRGB"] = dxf.getElementsByTagName("font")[0].getElementsByTagName("color")[0].getAttribute("rgb")
                        if dxf.getElementsByTagName("fill"):
                            pFill = dxf.getElementsByTagName("fill")[0].getElementsByTagName("patternFill")[0]
                            dxfDic["bgColorRGB"] = pFill.getElementsByTagName("bgColor")[0].getAttribute("rgb")
                    tempDic["dxfId"] = dxfDic
                    dic[ref] = tempDic                
                
        cfDic["sheet{0}".format(i+1)] = dic
    return cfDic                  

def worksheetProtection(worksheets):
    protectDic = {}
    for i,worksheet in enumerate(worksheets):
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("sheetProtection"):
            protect = wsheet.getElementsByTagName("sheetProtection")[0].getAttribute("sheet")
            if protect == '1':
                res = True
            else:
                res = False
        else:
            res = False
        protectDic["sheet{0}".format(i+1)] = res
    return protectDic

def worksheetOrientation(worksheets):
    orientation = []
    for worksheet in worksheets:
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("pageSetup"):
            orientation.append(wsheet.getElementsByTagName("pageSetup")[-1].getAttribute("orientation")) # to avoid getting customView
        else:
            orientation.append("NA")
    return orientation
      
def worksheetDataValidation(worksheets):
    dvDic = {}
    for i,worksheet in enumerate(worksheets):
        dic = {}
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("dataValidations"):
            DVs = wsheet.getElementsByTagName("dataValidations")[0]
            for dv in DVs.getElementsByTagName("dataValidation"):
                tempDic = {}
                ref = dv.getAttribute("sqref")
                tempDic["type"] = dv.getAttribute("type")
                if dv.hasAttribute("operator"):
                    tempDic["op"] = dv.getAttribute("operator")
                if dv.getElementsByTagName("formula1"):
                    tempDic["formula1"] = dv.getElementsByTagName("formula1")[0].firstChild.data
                if dv.getElementsByTagName("formula2"):
                    tempDic["formula2"] = dv.getElementsByTagName("formula2")[0].firstChild.data
                dic[ref] = tempDic                
        dvDic["sheet{0}".format(i+1)] = dic
    return dvDic
 
def worksheetsWithDrawing(worksheets): # I can check if it has a drawing or not, but I cannot know how many drawing it has
    wsList = []
    for i,worksheet in enumerate(worksheets):
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("drawing"):
            wsList.append("sheet{0}".format(i+1))
    return wsList

def worksheetMergeCells(worksheets):
    mcDic = {}
    for i,worksheet in enumerate(worksheets):
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("mergeCells"):
            mergeCells = wsheet.getElementsByTagName("mergeCells")[0]
            mcList = []
            for mc in mergeCells.getElementsByTagName("mergeCell"):
                mcList.append(mc.getAttribute("ref"))
            mcDic["sheet{0}".format(i+1)] = mcList
    return mcDic

def worksheetRowHeight(wsheet,r): # r is the row number we check the height for
    sheet = accessWorksheet(wsheet)
    if sheet:
        if sheet.getElementsByTagName("sheetData"):
            sheetData = sheet.getElementsByTagName("sheetData")[0]
            for row in sheetData.getElementsByTagName("row"):
                if int(row.getAttribute("r")) == r:
                    if row.hasAttribute("ht"):
                        return float(row.getAttribute("ht"))
                    
                    
def getHeaderFooter(worksheet,hf):
    wsheet = accessWorksheet(worksheet)
    if wsheet.getElementsByTagName("headerFooter"):
        header = None
        footer = None
        headFoot = wsheet.getElementsByTagName("headerFooter")[0]
        if headFoot.getElementsByTagName("oddHeader") or headFoot.getElementsByTagName("oddFooter"):
            if headFoot.getElementsByTagName("oddHeader"):
                header = headFoot.getElementsByTagName("oddHeader")[0].firstChild.data.lower()
            if headFoot.getElementsByTagName("oddFooter"):
                footer = headFoot.getElementsByTagName("oddFooter")[0].firstChild.data.lower()
        else:
            print("Check how the header or footer is defined in the XMLs")
        if hf == 'header':
            return header
        if hf == 'footer':
            return footer
        print("The label should be header or footer")
        return np.nan
    
def headerFooterRe(worksheet,hf):
    txt = getHeaderFooter(worksheet,hf)
    left = None
    center = None
    right = None
    if txt:
        if '&l' in txt:
            r = re.search('&l(.*)&c',txt)
            if r:
                res = r.group(1)
            else:
                r = re.search('&l(.*)&r',txt)            
                if r:
                    res = r.group(1)
                else:
                    res = txt.split('&l')[1]
            left = res
        if '&c' in txt:
            r = re.search('&c(.*)&r',txt)            
            if r:
                res = r.group(1)
            else:
                res = txt.split('&c')[1]
            center = res
        if '&r' in txt:
            right = txt.split('&r')[1]
    return left,center,right


def containsTable(worksheet): # return how many tables are in the worksheet
    c = 0
    wsheet = accessWorksheet(worksheet)
    if wsheet.getElementsByTagName("tableParts"):
        tables = wsheet.getElementsByTagName("tableParts")[0]
        for table in tables.getElementsByTagName("tablePart"):
            c += 1
    return c

def worksheetWithTables(worksheets,tables):
    dic = {}
    if tables != []:        
        tblRef = [] # table references obtained from table files
        i = 0
        for tbl in tables:
            tblRef.append(tableReference(tbl))

        for w,worksheet in enumerate(worksheets):            
            if containsTable(worksheet) != 0:
                temp = []                
                for c in range(int(containsTable(worksheet))):                    
                    temp.append(tblRef[i+c])
                dic["sheet{0}".format(w+1)] = temp
                i += int(containsTable(worksheet))
    return dic
    
def worksheetPrint(wsheet): # works for printing in 1 page (check for both width and height)
    res = "WH" # has set up for both width and height
    if wsheet.getElementsByTagName("pageSetup"):
        pSetup = wsheet.getElementsByTagName("pageSetup")[0]
        if pSetup.hasAttribute("fitToHeight"):
            res = "W"
        if pSetup.hasAttribute("fitToWidth"):
            res = "H"
    return res
    
def worksheetsPrint(worksheets):
    printP = []
    for worksheet in worksheets:
        wsheet = accessWorksheet(worksheet)
        if wsheet.getElementsByTagName("sheetPr"):
            if wsheet.getElementsByTagName("sheetPr")[0].getElementsByTagName("pageSetUpPr"):
                pageSetUp = wsheet.getElementsByTagName("sheetPr")[0].getElementsByTagName("pageSetUpPr")[0]
                if pageSetUp.hasAttribute("fitToPage"):
                    if pageSetUp.getAttribute("fitToPage") == '1':
                        printP.append(worksheetPrint(wsheet))
                    else:
                        printP.append("NA")
                else:
                    printP.append("NA")
            else:
                printP.append("NA")
        else:
            printP.append("NA")
    return printP


# ### Charts

# In[59]:


def accessChart(chartFile):
    if chartFile:
        if chartFile.getElementsByTagName("c:chartSpace"):
            chartSpace = chartFile.getElementsByTagName("c:chartSpace")[0]
            if chartSpace.getElementsByTagName("c:chart"):
                return chartSpace.getElementsByTagName("c:chart")[0]
            
def getChartTitle(ctitle):
    title = ''
    if ctitle:
        if ctitle.getElementsByTagName("c:tx"):
            ctx = ctitle.getElementsByTagName("c:tx")[0]            
            c = 0
            for ap in ctx.getElementsByTagName("a:p"):
                if ap.getElementsByTagName("a:r"):
                    for ar in ap.getElementsByTagName("a:r"):
                        title += ar.getElementsByTagName("a:t")[0].firstChild.data.lower()
                if c >= 1:
                    title += "\n"    
                c += 1
            if title == '':
                if ctx.getElementsByTagName("c:v"):
                    if ctx.getElementsByTagName("c:v")[0].firstChild:
                        title = ctx.getElementsByTagName("c:v")[0].firstChild.data.lower()
    return title      
        
def getBarChart(barChart):
    dic = {}
    if barChart.getElementsByTagName("c:barDir"):
        if barChart.getElementsByTagName("c:barDir")[0].hasAttribute("val"):
            dic["type"] = barChart.getElementsByTagName("c:barDir")[0].getAttribute("val") 
    if barChart.getElementsByTagName("c:grouping"):
        if barChart.getElementsByTagName("c:grouping")[0].hasAttribute("val"):
            dic["group"] = barChart.getElementsByTagName("c:grouping")[0].getAttribute("val") 
    if barChart.getElementsByTagName("c:ser"):
        color = []
        ref1 = []
        ref2 = []
        for ser in barChart.getElementsByTagName("c:ser"):
            if ser.getElementsByTagName("c:spPr"):
                if ser.getElementsByTagName("c:spPr")[-1].getElementsByTagName("a:solidFill"): # sometimes it's in the dPt (-1)
                    fill = ser.getElementsByTagName("c:spPr")[-1].getElementsByTagName("a:solidFill")[0]
                    if fill.getElementsByTagName("a:schemeClr"):
                        color.append(fill.getElementsByTagName("a:schemeClr")[0].getAttribute("val"))
                    elif fill.getElementsByTagName("a:srgbClr"):
                        color.append(fill.getElementsByTagName("a:srgbClr")[0].getAttribute("val"))
                    else:
                        print("The solid color is not defined with either schemeClr or srgbClr")
            if ser.getElementsByTagName("c:cat"):
                cat = ser.getElementsByTagName("c:cat")[0]
                if cat.getElementsByTagName("c:f"):
                    ref1.append(cat.getElementsByTagName("c:f")[0].firstChild.data.lower())
                    # here we can add actual categories if needed
            if ser.getElementsByTagName("c:tx"):
                tx = ser.getElementsByTagName("c:tx")[0]
                if tx.getElementsByTagName("c:f"):
                    ref1.append(tx.getElementsByTagName("c:f")[0].firstChild.data.lower())
            if ser.getElementsByTagName("c:val"):
                val = ser.getElementsByTagName("c:val")[0]
                if val.getElementsByTagName("c:numRef"):
                    numRef = val.getElementsByTagName("c:numRef")[0]
                    if numRef.getElementsByTagName("c:f"):
                        ref2.append(numRef.getElementsByTagName("c:f")[0].firstChild.data.lower())
        dic["color"] = color
        dic["ref1"] = ref1
        dic["ref2"] = ref2
    return dic

def getLineChart(lineChart):
    dic = {}
    dic["type"] = "line"
    if lineChart.getElementsByTagName("c:grouping"):
        if lineChart.getElementsByTagName("c:grouping")[0].hasAttribute("val"):
            dic["group"] = lineChart.getElementsByTagName("c:grouping")[0].getAttribute("val") 
    if lineChart.getElementsByTagName("c:ser"):
        color = []
        ref = []
        for ser in lineChart.getElementsByTagName("c:ser"):
            if ser.getElementsByTagName("c:spPr"):
                if ser.getElementsByTagName("c:spPr")[-1].getElementsByTagName("a:solidFill"): # sometimes it's in the dPt (-1)
                    fill = ser.getElementsByTagName("c:spPr")[-1].getElementsByTagName("a:solidFill")[0]
                    if fill.getElementsByTagName("a:schemeClr"):
                        color.append(fill.getElementsByTagName("a:schemeClr")[0].getAttribute("val"))
                    elif fill.getElementsByTagName("a:srgbClr"):
                        color.append(fill.getElementsByTagName("a:srgbClr")[0].getAttribute("val"))
                    else:
                        print("The solid color is not defined with either schemeClr or srgbClr")

            if ser.getElementsByTagName("c:val"):
                val = ser.getElementsByTagName("c:val")[0]
                if val.getElementsByTagName("c:numRef"):
                    numRef = val.getElementsByTagName("c:numRef")[0]
                    if numRef.getElementsByTagName("c:f"):
                        ref.append(numRef.getElementsByTagName("c:f")[0].firstChild.data.lower())
                else:
                    print("The reference is not defined through numRef")
            else:
                print("The reference is not defined through val")
        dic["color"] = color
        dic["ref"] = ref
    return dic

def getAxesNo(plotArea): # number of vertical axes
    c = 0
    if plotArea:        
        for axis in plotArea.getElementsByTagName("c:valAx"):
            c += 1
    return c
                
def getTwoCharts(barChart,lineChart,axes):
    dic = {}
    dic["type"] = "twoCharts"
    dic["bar"] = barChart
    dic["line"] = lineChart
    dic["axes"] = axes
    return dic
    
        
def getChart(chartFile):
    flag = 0
    chart = accessChart(chartFile)
    if chart:
        chartDic = {}
        
        title = ''
        if chart.getElementsByTagName("c:title"):
            title = getChartTitle(chart.getElementsByTagName("c:title")[0])
        
        
        theChart = None
        if chart.getElementsByTagName("c:plotArea"):
            plotArea = chart.getElementsByTagName("c:plotArea")[0]
            if not title: # for cases that the title is written in the plot area
                title = getChartTitle(plotArea)
            if plotArea.getElementsByTagName("c:barChart"):
                flag = 1
                theChart = getBarChart(plotArea.getElementsByTagName("c:barChart")[0])
            if plotArea.getElementsByTagName("c:bar3DChart"):
                theChart = getBarChart(plotArea.getElementsByTagName("c:bar3DChart")[0])
            if plotArea.getElementsByTagName("c:lineChart"):
                if flag == 1:
                    axes = getAxesNo(plotArea)
                    theChart = getTwoCharts(theChart,getLineChart(plotArea.getElementsByTagName("c:lineChart")[0]),axes)
                else:
                    theChart = getLineChart(plotArea.getElementsByTagName("c:lineChart")[0])
                
        chartDic["title"] = title
        if theChart:
            chartDic.update(theChart)
        return chartDic
    
            
def getChartStartCell(xdr):
    if xdr.getElementsByTagName("xdr:from"):
        dic = {}
        fromCR = xdr.getElementsByTagName("xdr:from")[0]
        dic["fromCol"] = int(fromCR.getElementsByTagName("xdr:col")[0].firstChild.data)                                     
        dic["fromRow"] = int(fromCR.getElementsByTagName("xdr:row")[0].firstChild.data) 
        return dic
    
    
def getPicture(xdr):
    if xdr:
        dic = getChartStartCell(xdr)
        if not dic:
            print("the drawing does not have an indication for its start cell")
            dic = {}
        if xdr.getElementsByTagName("xdr:pic"):
            pic = xdr.getElementsByTagName("xdr:pic")[0]
            if pic.getElementsByTagName("xdr:spPr"):
                aext = pic.getElementsByTagName("xdr:spPr")[0].getElementsByTagName("a:ext")[0]
                dic["cx"] = aext.getAttribute("cx")
                dic["cy"] = aext.getAttribute("cy")
            if pic.getElementsByTagName("a:duotone"):
                duotone = pic.getElementsByTagName("a:duotone")[0]
                tempDic = {}
                tempDic["prstClr"] = duotone.getElementsByTagName("a:prstClr")[0].getAttribute("val")
                if duotone.getElementsByTagName("a:schemeClr"):
                    tempDic["schemeClr"] = duotone.getElementsByTagName("a:schemeClr")[0].getAttribute("val")
                if duotone.getElementsByTagName("a:srgbClr"):
                    tempDic["srgbClr"] = duotone.getElementsByTagName("a:srgbClr")[0].getAttribute("val")
                dic["recolor"] = tempDic
            else:
                dic["recolor"] = "No recolor"
            if pic.getElementsByTagName("a14:imgEffect"):
                dic["imgEffect"] = str(pic.getElementsByTagName("a14:imgEffect")[0].firstChild).split()[2]
            return dic
    
                            
def drawingFile(drawingFiles,worksheets,chartFiles): # pictures in a worksheet are in the order of their creation
    wsList = worksheetsWithDrawing(worksheets)
    drawingDic = {}
    cC = 0 # counter on charts in a drawing XML file
    for i,drawingXML in enumerate(drawingFiles):
        drawings = {}
        charts = {}
        cP = 1 # counter on pictures in a drawing XML file        
        wsDr = drawingXML.getElementsByTagName("xdr:wsDr")[0]
        if wsDr.getElementsByTagName("xdr:oneCellAnchor"):
            for xdr in wsDr.getElementsByTagName("xdr:oneCellAnchor"): # for all pictures that described as oneAnchor
                dic = getPicture(xdr)
                if dic:
                    if drawings.get("picture1","empty") == "empty":
                        drawings["picture1"] = dic
                    else:
                        cP += 1
                        print("there is more than one picture in",wsList[i])
                        drawings["picture{0}".format(cP)] = dic
        if wsDr.getElementsByTagName("xdr:twoCellAnchor"):
            for xdr in wsDr.getElementsByTagName("xdr:twoCellAnchor"): # charts or pictures
                if xdr.hasAttribute("editAs"):
                    if xdr.getAttribute("editAs") == "oneCell": # it is a picture
                        if xdr.getElementsByTagName("xdr:pic"):
                            dic = getPicture(xdr)

                            if drawings.get("picture1","empty") == "empty":
                                drawings["picture1"] = dic
                            else:
                                cP += 1
                                print("there is more than one picture in",wsList[i])
                                drawings["picture{0}".format(cP)] = dic
                        else:
                            print("No Pic is available!")
                            
                    elif xdr.getAttribute("editAs") == "absolute": # it is a chart
                        dic = getChartStartCell(xdr)
                        if not dic:
                            print("the drawing does not have an indication for its start cell")
                            dic = {}     
                        dic.update(getChart(chartFiles[cC]))
                        charts["chart{0}".format(cC+1)] = dic
                        cC += 1
                    else:
                        print("There is a type other than chart or picture in the drawing.")

                else: # charts
                    dic = getChartStartCell(xdr)
                    if not dic:
                        print("the drawing does not have an indication for its start cell")
                        dic = {}     
                    dic.update(getChart(chartFiles[cC]))
                    charts["chart{0}".format(cC+1)] = dic
                    cC += 1
                
        drawingDic[wsList[i]] = drawings
        drawingDic[wsList[i]].update(charts)
    return drawingDic
          
        
# def chartFile(chartXML):
#     chartSpace = chartXML.getElementsByTagName("c:chartSpace")[0]
#     cChart = chartSpace.getElementsByTagName("c:chart")[0]    
#     plotArea = cChart.getElementsByTagName("c:plotArea")[0]
#     ctitle = None
#     if cChart.getElementsByTagName("c:title"):
#         ctitle = cChart.getElementsByTagName("c:title")[0]
#     chartDic = {}
#     title = ""
#     chart = None
#     if ctitle:
#         if ctitle.getElementsByTagName("c:tx"):
#             ctx = ctitle.getElementsByTagName("c:tx")[0]
#             crich = ctx.getElementsByTagName("c:rich")[0]
#             c = 0
#             for ap in crich.getElementsByTagName("a:p"):
#                 if ap.getElementsByTagName("a:r"):
#                     for ar in ap.getElementsByTagName("a:r"):
#                         title += ar.getElementsByTagName("a:t")[0].firstChild.data
#                 if c >= 1:
#                     title += "\n"    
#                 c += 1
        
#     if plotArea.getElementsByTagName("c:lineChart"):
#         chart = plotArea.getElementsByTagName("c:lineChart")[0]
#         chartDic["type"] = "lineChart"

#     elif plotArea.getElementsByTagName("c:barChart"):
#         chart = plotArea.getElementsByTagName("c:barChart")[0]
#         chartDic["type"] = "barChart"
        
#     if chart:       
#         if chart.getElementsByTagName("c:grouping"):
#             grouping = chart.getElementsByTagName("c:grouping")[0].getAttribute("val")
#             chartDic["group"] = grouping
#         else:
#             print("There was no grouping value in the chart")
#         if chart.getElementsByTagName("c:ser"):
#             c = 0
#             for ser in chart.getElementsByTagName("c:ser"):
#                 # if the theme is correct, accents are enough to indicate colors
#                 if ser.getElementsByTagName("c:spPr")[0].getElementsByTagName("a:solidFill"):
#                     lineColor = ser.getElementsByTagName("c:spPr")[0].getElementsByTagName("a:solidFill")[0]
#                     if lineColor.getElementsByTagName("a:schemeClr"):
#                         color = lineColor.getElementsByTagName("a:schemeClr")[0].getAttribute("val")
#                     elif lineColor.getElementsByTagName("a:srgbClr"):
#                         # Here we can add transparency if needed
#                         color = lineColor.getElementsByTagName("a:srgbClr")[0].getAttribute("val")
#                     else:
#                         print("Color in the Chart is not defined by scheme or RGB colors for the Solid pattern")
#                 if ser.getElementsByTagName("c:spPr")[0].getElementsByTagName("a:pattFill"):
#                     lineColor = ser.getElementsByTagName("c:spPr")[0].getElementsByTagName("a:pattFill")[0]
#                     fgColor = lineColor.getElementsByTagName("a:fgClr")[0].getElementsByTagName("a:schemeClr")[0].getAttribute("val")
#                     bgColor = lineColor.getElementsByTagName("a:bgClr")[0].getElementsByTagName("a:schemeClr")[0].getAttribute("val")
#                     color = (fgColor,bgColor)

#                 chartDic["color{0}".format(c)] = color
#                 ref = ser.getElementsByTagName("c:val")[0].getElementsByTagName("c:numRef")[0]\
#                 .getElementsByTagName("c:f")[0].firstChild.data
#                 chartDic["ref{0}".format(c)] = ref
#                 c += 1                
#         else:
#             print("There was no ser value in the line chart")
     
#     else:
#         if plotArea.getElementsByTagName("c:scatterChart"):
#             chart = plotArea.getElementsByTagName("c:scatterChart")[0]
#             chartDic["type"] = "scatterChart"
#             if chart.getElementsByTagName("c:ser"):
#                 ser = chart.getElementsByTagName("c:ser")[0]
#                 if title == '':
#                     title = ser.getElementsByTagName("c:tx")[0].getElementsByTagName("c:strRef")[0].\
#                     getElementsByTagName("c:strCache")[0].getElementsByTagName("c:pt")[0].getElementsByTagName("c:v")[0].\
#                     firstChild.data
#                 # this is the Fill. The other color is the outline that I didn't assign
#                 fill = ser.getElementsByTagName("c:marker")[0].getElementsByTagName("c:spPr")[0].\
#                 getElementsByTagName("a:solidFill")[0]
#                 if fill.getElementsByTagName("a:srgbClr"):
#                     color = fill.getElementsByTagName("a:srgbClr")[0].getAttribute("val")
#                 elif fill.getElementsByTagName("a:schemeClr"):
#                     color = fill.getElementsByTagName("a:schemeClr")[0].getAttribute("val")
#                 chartDic["color0"] = color
#                 ref0 = ser.getElementsByTagName("c:xVal")[0].getElementsByTagName("c:numRef")[0].\
#                 getElementsByTagName("c:f")[0].firstChild.data
#                 ref1 = ser.getElementsByTagName("c:yVal")[0].getElementsByTagName("c:numRef")[0].\
#                 getElementsByTagName("c:f")[0].firstChild.data
#                 chartDic["ref0"] = ref0
#                 chartDic["ref1"] = ref1

#         elif plotArea.getElementsByTagName("c:pieChart"):
#             chart = plotArea.getElementsByTagName("c:pieChart")[0]
#             chartDic["type"] = "pieChart"
#             if chart.getElementsByTagName("c:ser"):
#                 ser = chart.getElementsByTagName("c:ser")[0]
#                 c = 0
#                 for dpt in ser.getElementsByTagName("c:dPt"):
#                     sppr = dpt.getElementsByTagName("c:spPr")[0]
#                     if sppr.getElementsByTagName("a:solidFill"):
#                         color = sppr.getElementsByTagName("a:solidFill")[0].getElementsByTagName("a:schemeClr")[0].getAttribute("val")
#                     # it works as there is no pattFill when we have solid colors
#                     if sppr.getElementsByTagName("a:pattFill"):
#                         patt = sppr.getElementsByTagName("a:pattFill")[0]
#                         fgColor = patt.getElementsByTagName("a:fgClr")[0].getElementsByTagName("a:schemeClr")[0].getAttribute("val")
#                         bgColor = patt.getElementsByTagName("a:bgClr")[0].getElementsByTagName("a:schemeClr")[0].getAttribute("val")
#                         color = (fgColor,bgColor)
#                     chartDic["color{0}".format(c)] = color
#                     c += 1
#                 strRef = ser.getElementsByTagName("c:cat")[0].getElementsByTagName("c:strRef")[0]
#                 ref0 = strRef.getElementsByTagName("c:f")[0].firstChild.data
#                 chartDic["ref0"] = ref0
#                 strCache = strRef.getElementsByTagName("c:strCache")[0]
#                 c = 0
#                 for pt in strCache.getElementsByTagName("c:pt"):
#                     chartDic["label{0}".format(c)] = pt.getElementsByTagName("c:v")[0].firstChild.data  
#                     c += 1
                
#         else:
#             print("No chart has been loaded")
#     chartDic["title"] = title
#     return chartDic

# def assignChart(chartFiles,drawings):
#     i = 0
#     chartsDic = {}
#     for sheet,drawing in drawings.items():
#         c = 0
#         for item in drawing:
#             if item != "picture": # the item is a chart
#                 c += 1
#                 dic = {}
#                 dic = chartFile(chartFiles[i])
#                 dic["size"] = item
#                 chartsDic[sheet] = dic
#                 i += 1
#             if c > 1:
#                 print(sheet," has more than one chart in it.")
#     return chartsDic


# ### Pivot Tables 

# In[60]:


def accessPivotTable(pTableFile):
    if pTableFile:
        if pTableFile.getElementsByTagName("pivotTableDefinition"):
            return pTableFile.getElementsByTagName("pivotTableDefinition")[0]
        
def getPivotCache(pcFile):
    if pcFile:
        res = [] # cache fields in the pivotCacheFile
        pivotCache = pcFile.getElementsByTagName("pivotCacheDefinition")[0]
        if pivotCache.getElementsByTagName("cacheFields"):
            cacheFields = pivotCache.getElementsByTagName("cacheFields")[0]
            for field in cacheFields.getElementsByTagName("cacheField"):
                if field.hasAttribute("name"):
                    res.append(field.getAttribute("name").lower())
        return res
    
def getCacheId(pTableFile):
    if accessPivotTable(pTableFile):
        table = accessPivotTable(pTableFile)
        if table.hasAttribute("cacheId"):
            return table.getAttribute("cacheId")
    
def pivotCaches(pCacheFiles):
    dic = {}
    for i,file in enumerate(pCacheFiles):
        dic[i] = getPivotCache(file)
    return dic
    
def getPivotAreas(pivotTable,cacheList):
    pvTable = {}
    if pivotTable.getElementsByTagName("rowFields"):
        temp = []
        rowFields = pivotTable.getElementsByTagName("rowFields")[0]
        for rField in rowFields.getElementsByTagName("field"):
            ind = int(rField.getAttribute("x"))
            if ind < 0: # e.g. using values as rows
                temp.append(str(ind))
            elif len(cacheList)>ind:
                temp.append(cacheList[ind])
            else:
                print("This should not happen! Wrong cache list is fed. Check row fields")
        pvTable["rows"] = temp
                
    if pivotTable.getElementsByTagName("colFields"):
        temp = []
        colFields = pivotTable.getElementsByTagName("colFields")[0]
        for cField in colFields.getElementsByTagName("field"):
            ind = int(cField.getAttribute("x"))
            if ind < 0: # e.g. using values as rows
                temp.append(str(ind))
            elif len(cacheList)>ind:
                temp.append(cacheList[ind])
            else:
                print("This should not happen! Wrong cache list is fed. Check column fields")  
        pvTable["columns"] = temp
        
    if pivotTable.getElementsByTagName("pageFields"):
        temp = []
        pageFields = pivotTable.getElementsByTagName("pageFields")[0]
        for pField in pageFields.getElementsByTagName("pageField"):
            ind = int(pField.getAttribute("fld"))
            if ind < 0: # e.g. using values as rows
                temp.append(str(ind))
            elif len(cacheList)>ind:
                temp.append(cacheList[ind])
            else:
                print("This should not happen! Wrong cache list is fed. Check filters")  
        pvTable["filters"] = temp

    if pivotTable.getElementsByTagName("dataFields"):
        temp = []
        dataFields = pivotTable.getElementsByTagName("dataFields")[0]
        for dField in dataFields.getElementsByTagName("dataField"):
            name = dField.getAttribute("name").lower()
            show = ''
            if dField.hasAttribute("showDataAs"):
                show = dField.getAttribute("showDataAs").lower()
            temp.append(name+' '+show)
        pvTable["values"] = temp
    
    return pvTable
        
def getPivotRef(pivotTable):
    if pivotTable.getElementsByTagName("location"):
        location = pivotTable.getElementsByTagName("location")[0]
        if location.hasAttribute("ref"):
            return location.getAttribute("ref")
        
def getCacheInd(cacheId,sortedId,i):
    if len(cacheId)>i:
        temp = cacheId[i]
        for j,Id in enumerate(sortedId):
            if Id == temp:
                return j
    else:
        print("This should not happen. check cacheId list")    
                
def getPivotTables(pTableFiles,pCacheFiles): # I think the smallest cacheId related to the first cache xml file
    pt = {}
    cacheList = pivotCaches(pCacheFiles)
    cId = [] # cacheId
    for file in pTableFiles:
        cId.append(int(getCacheId(file)))
    sortId = sorted(set(cId))
    for i,table in enumerate(pTableFiles):
        ind = getCacheInd(cId,sortId,i)
        tableObj = accessPivotTable(table)
        ref = getPivotRef(tableObj)
        pt[ref] = getPivotAreas(tableObj,cacheList[ind])
    return pt


# ### Compare Functions

# In[61]:


#################################################################################################################################
# This code is not written by me! I have copied it from here: http://norvig.com/spell-correct.html
# The pyEnchant Library has many problems, but this peice of code is enough for our purpose
import re
from collections import Counter

def words(text): return re.findall(r'\w+', text.lower())

WORDS = Counter(words(open('big.txt').read()))

def P(word, N=sum(WORDS.values())): 
    "Probability of `word`."
    return WORDS[word] / N

def correction(word): 
    "Most probable spelling correction for word."
    return max(candidates(word), key=P)

def candidates(word): 
    "Generate possible spelling corrections for word."
    return (known([word]) or known(edits1(word)) or known(edits2(word)) or [word])

def known(words): 
    "The subset of `words` that appear in the dictionary of WORDS."
    return set(w for w in words if w in WORDS)

def edits1(word):
    "All edits that are one edit away from `word`."
    letters    = 'abcdefghijklmnopqrstuvwxyz'
    splits     = [(word[:i], word[i:])    for i in range(len(word) + 1)]
    deletes    = [L + R[1:]               for L, R in splits if R]
    transposes = [L + R[1] + R[0] + R[2:] for L, R in splits if len(R)>1]
    replaces   = [L + c + R[1:]           for L, R in splits if R for c in letters]
    inserts    = [L + c + R               for L, R in splits for c in letters]
    return set(deletes + transposes + replaces + inserts)

def edits2(word): 
    "All edits that are two edits away from `word`."
    return (e2 for e1 in edits1(word) for e2 in edits1(e1))
#################################################################################################################################


# In[62]:


def similar(a,b):
    return SequenceMatcher(None, a, b).ratio()

def previousCell(cellRef): # previous cell reference in a row
    cell = findCell(cellRef)
    if len(cell[0]) == 1:
        return chr(ord(cell[0])-1)+str(cell[1])
    else:
        return cell[0][0]+chr(ord(cell[0][-1])-1)+str(cell[1])
    
def nextCell(cellRef): # in a row
    cell = findCell(cellRef)
    if len(cell[0]) == 1:
        return chr(ord(cell[0])+1)+str(cell[1])
    else:
        return cell[0][0]+chr(ord(cell[0][-1])+1)+str(cell[1])
    
def nextCellCol(cellRef): # in a column
    cell = findCell(cellRef)
    return cell[0]+str(cell[1]+1)

def findRange(start,end): # on a column
    cells = [start]
    temp = start
    while temp != end:
        temp = nextCellCol(temp)
        cells.append(temp)        
    return cells

def compareNan(dic1,dic2):
    if dic1 == "empty":
        res = np.nan
    elif dic1 == dic2:
        res = True
    else:
        res = False
    return res


def compareAll(lstT,lstS):
    comp = []
    for wordT,wordS in zip(lstT,lstS):
        if wordT == correction(wordS) or wordT == correction(wordS)+'s' or wordT+'s' == correction(wordS)        or wordT == wordS+'s' or wordT == wordS:
            comp.append(True)
        else:
            comp.append(False)

    if all(item == True for item in comp):
        return True
    return False
        
def isfloat(value):
    try:
        float(value) or int(value)
        return True
    except ValueError:
        return False

    
def checkEither(dT,dS,v1,v2): # the exact match is already tested
    if set([dT,dS]) == set([v1,v2]):
        return True
    return False
    
def checkBoolean(dT,dS):
    if checkEither(dT,dS,"true","1"):
        return True
    elif checkEither(dT,dS,"false","0"):
        return True
    else:
        return False
    
def removeChar(word):
    regex = re.compile("[“”()'’:,\.!?]")
    word = regex.sub('', word)
    return word    

def checkData(dT,dS):
    if not (dT or dS): # both are None
        return np.nan # this case shouldn't happen
    elif not (dT and dS): # one of them is None
        return False
    else:
        dT = dT.lower()
        dS = dS.lower()
        if dT == dS:
            return True
        else:
            if isfloat(dT) and isfloat(dS):
                if abs(float(dT)-float(dS))<0.1:
                    return True
                else:
                    return False
            elif checkBoolean(dT,dS):
                return True            

            elif dT.replace(" ", "") == dS.replace(" ", ""):
                return True
            
            elif removeChar(dT) == removeChar(dS):
                return True
            
            elif len(dT.split())==len(dS.split()):
                res = compareAll(dT.split(),dS.split())
                return res

            else:        
                return False

            
def sameFormula(fT,fS,rangeOp,commaOp):
    if rangeOp in fT and commaOp in fS:
        sumCell = findCell(fT)
        plusCell = findCell(fS)
        plusCells = []
        for c in range(len(plusCell)):
            if c%2 == 0:
                plusCells.append(plusCell[c]+str(plusCell[c+1]))
                
        if "," in fT:
            if set(sumCell) == set(plusCell):
                return True
            else:
                return False
        elif ":" in fT:
            temp = []
            sumCells = []
            for i in range(sumCell[3]-sumCell[1]+1):
                for j in range(ord(sumCell[2])-ord(sumCell[0])+1):
                    sumCells.append(chr(ord(sumCell[0])+j)+str(sumCell[1]+i))
            if set(sumCells) == set(plusCells):
                return True
            else:
                return False
        else:
            print("This should not happen. sumFormula function.")
    elif rangeOp in fS and commaOp in fT:
        return sameFormula(fS,fT,rangeOp,commaOp)
    else:
        return False
       
def misspelling(fT,fS): # works if the sequence is correct
    regex = r'\b([a-z]+)\b'
    wordsT = re.findall(regex,fT)
    wordsS = re.findall(regex,fS)
    for wordT, wordS in zip(wordsT,wordsS):
        if similar(wordT,wordS)>.8 and similar(wordT,wordS) != 1:
            fS = fS.replace(wordS,wordT)
    if fS == fT:
        return True
    else:
        return False
    
def checkFormula(fT,fS):
    if not (fT or fS): # both are None
        return np.nan
    elif not (fT and fS): # one of them is None
        return False
    else:
        fT = fT.lower() 
        fS = fS.lower()
        if fT == fS: # the exact match
            return True
        elif fT.replace(" ", "") == fS.replace(" ", ""):
            return True
        else:
            if sameFormula(fT,fS,"sum","+"):
                return True
#             elif sameFormula(fT,fS,"concat","concatenate"):
#                 return True
#             elif sameFormula(fT,fS,"concat","textjoin"):
#                 return True
            elif misspelling(fT,fS):
                return True
        return np.nan   
            
        
def checkExactFormula(fT,fS):
    if not (fT or fS): # both are None
        return np.nan
    elif not (fT and fS): # one of them is None
        return False
    else:
        fT = fT.lower() 
        fS = fS.lower()
        if fT == fS: # the exact match
            return True
        elif fT.replace(" ", "") == fS.replace(" ", ""):
            return True
        else:            
            return False
        
def checkAlign(sT,sS):
    sDicT = stDicT[sT]
    sDicS = stDicS[sS]
    resList = []
    # align (default) OR alignH, alignV, alignI
    if sDicT.get("align","empty") == "empty" and sDicS.get("align","empty") == "empty":
        resList.append(np.nan)
    elif sDicT.get("align","empty") == sDicS.get("align","empty"): # they are both default or both have the same exact attributes
        resList.append(True)
    elif sDicT.get("align","empty") != "empty" and sDicS.get("align","empty") != "empty":
        sADicT = sDicT["align"]
        sADicS = sDicS["align"]
        resList.append(compareNan(sADicT.get("alignH","empty"),sADicS.get("alignH","empty")))
        resList.append(compareNan(sADicT.get("alignV","empty"),sADicS.get("alignV","empty")))
        resList.append(compareNan(sADicT.get("alignI","empty"),sADicS.get("alignI","empty")))
    else:
        resList.append(False)
    return resList
    
# I decided to only check the styles, as the colors were never asked
def checkBorder(sT,sS): # if the border styles is entirely correct, it returns True (boolean)
    sDicT = stDicT[sT]
    sDicS = stDicS[sS]
    resList = [] # style: left, right, top, bottom, daigonal    
    if sDicT.get("border","empty") == "empty":
        return np.nan
    elif sDicT.get("border","empty") == sDicS.get("border","empty"):
        return True
    elif sDicT.get("border","empty") != "empty" and sDicS.get("border","empty") != "empty":
        sBDicT = sDicT["border"]
        sBDicS = sDicS["border"]
        if sBDicT.get("left","empty") == sBDicS.get("left","empty"):
            resList.append(True)
        elif sBDicT.get("left","empty") != "empty" and sBDicS.get("left","empty") != "empty":
            tempDicT = sBDicT["left"]
            tempDicS = sBDicS["left"]
            resList.append(compareNan(tempDicT.get("style","empty"),tempDicS.get("style","empty"))) 
        else:
            resList.append(False)
        if sBDicT.get("right","empty") == sBDicS.get("right","empty"):
            resList.append(True)
        elif sBDicT.get("right","empty") != "empty" and sBDicS.get("right","empty") != "empty":
            tempDicT = sBDicT["right"]
            tempDicS = sBDicS["right"]
            resList.append(compareNan(tempDicT.get("style","empty"),tempDicS.get("style","empty"))) 
        else:
            resList.append(False)
        if sBDicT.get("top","empty") == sBDicS.get("top","empty"):
            resList.append(True)
        elif sBDicT.get("top","empty") != "empty" and sBDicS.get("top","empty") != "empty":
            tempDicT = sBDicT["top"]
            tempDicS = sBDicS["top"]
            resList.append(compareNan(tempDicT.get("style","empty"),tempDicS.get("style","empty"))) 
        else:
            resList.append(False)
        if sBDicT.get("bottom","empty") == sBDicS.get("bottom","empty"):
            resList.append(True)
        elif sBDicT.get("bottom","empty") != "empty" and sBDicS.get("bottom","empty") != "empty":
            tempDicT = sBDicT["bottom"]
            tempDicS = sBDicS["bottom"]
            resList.append(compareNan(tempDicT.get("style","empty"),tempDicS.get("style","empty")))
        else:
            resList.append(False)
        if sBDicT.get("diagonal","empty") == sBDicS.get("diagonal","empty"):
            resList.append(True)
        elif sBDicT.get("diagonal","empty") != "empty" and sBDicS.get("diagonal","empty") != "empty":
            tempDicT = sBDicT["diagonal"]
            tempDicS = sBDicS["diagonal"]
            resList.append(compareNan(tempDicT.get("style","empty"),tempDicS.get("style","empty")))  
        else:
            resList.append(False)
    else:
        resList.append(False)
    if all(item == True for item in resList):
        return True
    return False      
    

def checkFill(sT,sS):
    sDicT = stDicT[sT]
    sDicS = stDicS[sS]
    resList = [] # pattern, fgRGB, fgTheme, fgTint, bgIndex
    
    if sDicT.get("fill","empty") == "empty" and sDicS.get("fill","empty") == "empty":
        resList.append(np.nan)
    elif sDicT.get("fill","empty") == sDicS.get("fill","empty"):
        resList.append(True)
    elif sDicT.get("fill","empty") != "empty" and sDicS.get("fill","empty") != "empty":
        sFDicT = sDicT["fill"]
        sFDicS = sDicS["fill"]
        resList.append(compareNan(sFDicT.get("pattern","empty"),sFDicS.get("pattern","empty")))
        resList.append(compareNan(sFDicT.get("fgColorRGB","empty"),sFDicS.get("fgColorRGB","empty")))
        resList.append(compareNan(sFDicT.get("fgColorTheme","empty"),sFDicS.get("fgColorTheme","empty")))
        resList.append(compareNan(sFDicT.get("fgColorTint","empty"),sFDicS.get("fgColorTint","empty")))
        resList.append(compareNan(sFDicT.get("bgColorI","empty"),sFDicS.get("bgColorI","empty")))
    else:
        resList.append(False)
    return resList  


def checkFont(sT,sS):
    sDicT = stDicT[sT]
    sDicS = stDicS[sS]
    resList = []
    
    # bold, underline, italic, strike, size, colorT, colorRGB, name, family, scheme
    if sDicT.get("font","empty") == "empty" and sDicS.get("font","empty") == "empty":
        resList.append(np.nan)
    elif sDicT.get("font","empty") == sDicS.get("font","empty"):
        resList.append(True)
    elif sDicT.get("font","empty") != "empty" and sDicS.get("font","empty") != "empty":
        sFDicT = sDicT["font"]
        sFDicS = sDicS["font"]
        resList.append(compareNan(sFDicT.get("bold","empty"),sFDicS.get("bold","empty")))
        resList.append(compareNan(sFDicT.get("underline","empty"),sFDicS.get("underline","empty")))
        resList.append(compareNan(sFDicT.get("italic","empty"),sFDicS.get("italic","empty")))
        resList.append(compareNan(sFDicT.get("strike","empty"),sFDicS.get("strike","empty")))
        resList.append(compareNan(sFDicT.get("size","empty"),sFDicS.get("size","empty")))
        resList.append(compareNan(sFDicT.get("colorT","empty"),sFDicS.get("colorT","empty")))
        resList.append(compareNan(sFDicT.get("colorRGB","empty"),sFDicS.get("colorRGB","empty")))
        resList.append(compareNan(sFDicT.get("name","empty"),sFDicS.get("name","empty")))
        resList.append(compareNan(sFDicT.get("family","empty"),sFDicS.get("family","empty")))
        resList.append(compareNan(sFDicT.get("scheme","empty"),sFDicS.get("scheme","empty")))

    else:
        resList.append(False)
    return resList  

def checkNumFmt(sT,sS):
    sDicT = stDicT[sT]
    sDicS = stDicS[sS]
    if sDicT.get("numFmt","empty") == sDicS.get("numFmt","empty"): # if they are empty the number format is general
        return True
    if checkEither(sDicT.get("numFmt","empty"),sDicS.get("numFmt","empty"),"empty",'0'): # general numberFormatId is 0
        return True
    elif sDicT.get("numFmt","empty").isdigit() and (not sDicS.get("numFmt","empty").isdigit()) and    sDicS.get("numFmt","empty")!="empty":
        print("The student might have used a custom Number Format that has the same result in style #", sS)
        return False    
    return False
    
def checkCellProtection(sT,sS):
    sDicT = stDicT[sT]
    sDicS = stDicS[sS]
    return compareNan(sDicT.get("protect","empty"),sDicS.get("protect","empty"))
    
def checkSheetProtect(sheet):
    if sheetProtectT[sheet] == False:
        print("You probably have entered the sheet number incorrectly.",sheet,"is not protected in the teacher's spreadsheet")
        return False
    else:
        if sheetProtectS.get(sheet,"empty") == True:
            return True        
        return False        
    
def checkThemes(themesRef,themes):
    if themesRef == themes:
        return True
    else:
        return False

    
def checkSheetNames(sheetsRef,sheets):
    noSheets = len(sheets) # number of sheets
    i = 0
    resList = []
    for sheetR in sheetsRef:
        if i<noSheets:
            resList.append(checkData(sheetR,sheets[i]))
        else:
            resList.append(False)
        i += 1
    return resList
    
def checkColsW(sheet,cols=None):
    cols = len(colsWT[sheet]) if cols is None else cols # for when no cols number have passed
    resList = []
    colsT = colsWT[sheet]
    colsS = colsWS[sheet]
    colRef = {}
    for colInd in colsT: # to ignore the columns that have been created in the next question
        if colInd < cols:
            colRef[colInd] = colsT[colInd]
    for col in colRef:
        if colsS.get(col,"empty") == "bestFit":
            resList.append(True)
        elif colsS.get(col,"empty") == "empty":
            if float(colRef[col]) > 10:
                resList.append(False)
            else: # there was no need to resize
                resList.append(True)
        elif float(colsS.get(col,"empty")) >= float(colRef[col])-1:
            resList.append(True)
        else:
            resList.append(False)
    return resList

def checkColsHide(sheet):
    resList = []
    colsT = hidColsT[sheet]
    colsS = hidColsS[sheet]
    for col in colsT:
        if col in colsS: # the student won't be punished for extra columns that she hides
            resList.append(True)
        else:
            resList.append(False)
    return resList

def checkFreeze(sheet):
    fT = freezeT[sheet]
    fS = freezeS[sheet]
    if fT == {}:
        print("Warning! the teacher's",sheet,"does not have a frozen pane")
    if fT == fS:
        return True    
    return False

def containTextF(formula):
    formula = formula.lower()
    spliter = '"'
    if spliter in formula:
        theText = formula.split(spliter)[1]
        theText = theText.split(spliter)[0]
        return theText
    else:
        return formula

def checkDXFs(sheet,col):
    resList = [] # the type and formatting
    col = col.upper()
    if conditionalFT[sheet] == {} or conditionalFT[sheet].get(col,"empty") == "empty":
        resList.extend((np.nan,np.nan))
    else:
        if len(conditionalFS)<int(sheet[-1]):
            resList.extend((False,False))
        else:
            if conditionalFS[sheet] == {} or conditionalFS[sheet].get(col,"empty") == "empty":
                resList.extend((False,False))
            else:          
                cfT = conditionalFT[sheet][col].copy()
                cfS = conditionalFS[sheet][col].copy()
                # for the formatting
#                 cfT["dxfId"] = dxfDicT[int(conditionalFT[sheet][col]["dxfId"])]    
#                 cfS["dxfId"] = dxfDicS[int(conditionalFS[sheet][col]["dxfId"])]
                if cfT.get("bottom","empty")==cfS.get("bottom","empty") and cfT.get("percent","empty")==cfS.get("percent","empty"):
                    if cfT.get("type","empty")==cfS.get("type","empty") and cfT.get("operator","empty")==cfS.get("operator","empty")                    and cfT.get("formula","empty")==cfS.get("formula","empty"):
                        resList.append(True)
                    else:
                        # equal and containText sometimes result in the same thing
                        if checkEither(cfT.get("operator","empty"),cfS.get("operator","empty"),"containsText","equal") or                        (cfT.get("operator","empty") == "containsText" and cfS.get("operator","empty") == "containsText"):
                            if containTextF(cfT.get("formula","empty")) == containTextF(cfS.get("formula","empty")):
                                resList.append(True)
                            else:
                                resList.append(False)                                 
                        else:
                            resList.append(False)
                else:
                    resList.append(False)
                if cfT["dxfId"] == cfS["dxfId"]:
                    resList.append(True)
                else:
                    resList.append(False)
    return resList
        
def checkOrientation(sheet):
    n = int(sheet[-1])-1
    oT = worksheetOrientation(worksheetsT)
    oS = worksheetOrientation(worksheetsS)
    if oT[n] == "NA":
        print("You are probably asking the orientaion of the wrong page.")
        return np.nan
    elif len(oS) <= n:
        return False
    elif oT[n] == oS[n]:
        return True
    return False

def checkPrintPage(sheet):
    n = int(sheet[-1])-1
    pT = worksheetsPrint(worksheetsT)
    pS = worksheetsPrint(worksheetsS)
    if pT[n] == "NA":
        print("You are probably asking about the print preview of the wrong page.")
        return np.nan
    elif len(pS) <= n:
        return False
    elif pT[n] == pS[n]:
        return True
    return False
    
def generateBetweenCells(r):
    res = [r.split(':')[0]]
    theRange = findCell(r)
    if len(theRange) == 4:
        if theRange[0] == theRange[2]: 
            for i in range(abs(theRange[3]-theRange[1])):
                res.append(theRange[0]+str(i+theRange[1]+1))
        # on a row
        elif theRange[1] == theRange[3]: 
            for j in range(abs(ord(theRange[0].upper())-ord(theRange[2].upper()))):
                res.append(chr(ord(theRange[0].upper())+j+1)+str(theRange[1]))
        else:
            print("The student has used a rectangle range")
    return res

def checkDataValidation(sheet,column,row): # between has no operator
    cell = column.upper() + str(row)
    dvT = worksheetDataValidation(worksheetsT)
    dvS = worksheetDataValidation(worksheetsS)
    if dvT.get(sheet,"empty") == "empty":
        print("Teacher's spreadsheet does not have",sheet)
        return np.nan
    if dvT[sheet] == {}:
        print(sheet,"of the teacher's spreadsheet does not have a ny data validation")
        return np.nan
    if dvT[sheet].get(cell,"empty") == "empty":
        print("You are probably asking about DataValidation of the wrong cell.")
        return np.nan
    if dvS.get(sheet,"empty") == "empty":
        return False
    if dvS[sheet] == {}:
        return False
    if dvS[sheet].get(cell,"empty") != "empty":
        return dvT[sheet][cell] == dvS[sheet][cell]
    else: # if student defined the dataValidation on a range by mistake
        for key,value in dvS[sheet].items():
            if ':' in key:
                theRange = generateBetweenCells(key)
                if any(cell in cells for cells in theRange):
                    return dvT[sheet][cell] == value
    return False
    
def checkDefinedName(name): 
    defNamesT = definedNames(workbookT)
    if not defNamesT:
        print("There is no defined name in the teacher's document")
        return [np.nan]
    if defNamesT.get(name,"empty") == "empty":
        print("There is no defined name as",name,"exists in the teacher's document")
        return [np.nan]
    defNamesS = definedNames(workbookS)
    if defNamesS: # it is not None
        res = [] # a definedName with the "name" exists, the range is indicated correctly
        if defNamesS.get(name,"empty") != "empty":
            res.append(True)
            address = defNamesS[name].split('!')
            sheet = address[0]
            theRange = address[1]
            if checkData(defNamesT[name].split('!')[0],sheet) and defNamesT[name].split('!')[1] == theRange:
                res.append(True)
            else:
                res.append(False)
        else:
            res.append(False)
            for key,value in defNamesS.items():
                if checkData(defNamesT[name].split('!')[0],value.split('!')[0]) and defNamesT[name].split('!')[1] == value.split('!')[1]:
                    res.append(True)
            if len(res) == 1:
                res.append(False)
        return res
    return [False,False]
        
# *** incomplete, does work if there is only one merging in a sheet    
def checkMergeCells(sheet,ref=None): # depends on the sheet order. ref is used if we have more than one merge in a sheet
    mcsT = worksheetMergeCells(worksheetsT)
    if mcsT.get(sheet,"empty") == "empty":
        print("There is no merged cell in",sheet,"of the teacher's document")
        return np.nan
    mcsS = worksheetMergeCells(worksheetsS)
    if mcsS.get(sheet,"empty") == "empty":
        return False
    if not ref:
        if mcsT[sheet] == mcsS[sheet]:
            return True
#     else:
#         for refs in mcsT[sheet]:
    return False

def checkRowHeight(sheet,r):
    wsheetT = getWorksheet(worksheetsT,sheet)
    wsheetS = getWorksheet(worksheetsS,sheet)
    if not worksheetRowHeight(wsheetT,r):
        print("The",sheet,"of the teacher's document does not have a row",r)
        return np.nan
    if worksheetRowHeight(wsheetS,r):
        if abs(worksheetRowHeight(wsheetT,r)-worksheetRowHeight(wsheetS,r))<1:
            return True
    return False
   
def checkPictureFeature(picT,picS,feature):
    if feature == "width":
        return abs(float(picT["cx"])-float(picS.get("cx","empty")))<5000
    elif feature == "height":
        return abs(float(picT["cy"])-float(picS.get("cy","empty")))<5000
    elif feature == "size":
        return abs(float(picT["cx"])-float(picS.get("cx","empty")))<5000 and abs(float(picT["cy"])-float(picS.get("cy","empty")))<5000
    elif feature == "startCell":
        return picT["fromCol"] == picS["fromCol"] and picT["fromRow"] == picS["fromRow"]
    elif feature == "exist":
        return abs(picT["fromCol"]-picS["fromCol"])<3 and abs(picT["fromRow"]-picS["fromRow"])<3
    else:
        if picT.get(feature,"empty") == "empty":
            print("The picture in",sheet,"of the document does not have a feature",feature)
            return np.nan
        return picT[feature] == picS.get(feature,"empty")
    
    
def checkPicture(sheet,feature,pN=1):
    if drawingsT.get(sheet,"empty") == "empty":
        print("There is no drawing in",sheet,"of the teacher's document")
        return np.nan
    if drawingsT[sheet].get("picture"+str(pN),"empty") == "empty":
        print("There is no drawing in",sheet,"of the teacher's document")
        return np.nan
    if drawingsS.get(sheet,"empty") == "empty":
        return False
    if drawingsS[sheet].get("picture1","empty") == "empty": # there is no picture in that sheet of the student's spreadsheet
        return False
    picT = drawingsT[sheet]["picture"+str(pN)]
    picS = None
    if drawingsS[sheet].get("picture"+str(pN),"empty") != "empty":
        picS = drawingsS[sheet]["picture"+str(pN)]
        sub = pN
    else:
        sub = pN -1
        while sub > 0:
            if drawingsS[sheet].get("picture"+str(sub),"empty") == "empty":
                sub -= 1
            else:
                picS = drawingsS[sheet]["picture"+str(sub)]
                break            
    if picS:
        if feature == "inserted": # we cannot check that if a second picture is inserted! it always return true if one picture exists in the student document
            return True           # so in that case use exist
        if checkPictureFeature(picT,picS,feature) == True:
            return True
        if checkPictureFeature(picT,picS,feature) == False: # an instance can be there is a picture but it does not have the feature
            n = sub - 1
            while n > 0:
                picS = drawingsS[sheet]["picture"+str(n)]
                if checkPictureFeature(picT,picS,feature):
                    return True
                n -= 1
            return False
        return np.nan
    return False
    
    
def checkNoHyperlink(sheet): # teacher's document not needed
    sheetS = getWorksheet(worksheetsS,sheet)
    if sheetS.getElementsByTagName("hyperlinks"):
        return False
    return True

def checkHeaderFooter(sheet,hf,part):
    wsheetT = getWorksheet(worksheetsT,sheet)
    wsheetS = getWorksheet(worksheetsS,sheet)
    leftT,centerT,rightT = headerFooterRe(wsheetT,hf)
    if wsheetS:
        leftS,centerS,rightS = headerFooterRe(wsheetS,hf)
        if part == "left":
            if not leftT:
                print("The",part,"cell of the teacher's,",hf,"is empty")
                return np.nan
            return leftT == leftS
        if part == "right":
            if not rightT:
                print("The",part,"cell of the teacher's,",hf,"is empty")
                return np.nan
            return rightT == rightS
        if part == "center":
            if not centerT:
                print("The",part,"cell of the teacher's,",hf,"is empty")
                return np.nan
            return centerT == centerS
    return False

def checkWorksheetTable(sheet): # checks if a table is inserted (removed) in sheet
    wsheetTableT = worksheetWithTables(worksheetsT,tableFilesT)
    wsheetTableS = worksheetWithTables(worksheetsS,tableFilesS)
    if wsheetTableT.get(sheet,"empty") == "empty": # the table on that sheet is removed
        if wsheetTableS.get(sheet,"empty") == "empty":
            return True
        return False
    if wsheetTableS.get(sheet,"empty") != "empty":
        return True
    return False

def checkTableRef(sheet):
    if checkWorksheetTable(sheet) == True: # nan case won't go trough
        wsheetTableT = worksheetWithTables(worksheetsT,tableFilesT)
        wsheetTableS = worksheetWithTables(worksheetsS,tableFilesS)
        if wsheetTableT[sheet] == wsheetTableS[sheet]:
            return True
    return False

def getColNumber(col):
    col = col.lower()
    if len(col) == 1:
        return ord(col)-97
    else: # assuming the first letter is always 'a'
        return ord(col[-1])-71 # 26-97
    
def checkChart(sheet,col,row,twoCharts,twoChartsFeature): # title and axes are the common features for a combined chart
    if drawingsT.get(sheet,"empty") == "empty":
        print("There is no chart in",sheet,"of the teacher's spreadsheet")
        return np.nan
    chartT = None
    for key,value in drawingsT[sheet].items():
        if "chart" in key:
            if value["fromCol"] == getColNumber(col) and value["fromRow"] == row-1:
                chartT = value
    if not chartT:
        print("In",sheet,"of the teacher's spreadsheet, no chart is located at cell",col.upper(),row)
        return np.nan
    if drawingsS.get(sheet,"empty") == "empty":
        return False
    chartS = None
    flag = 0
    for key,value in drawingsS[sheet].items():
        if "chart" in key:
            flag = 1
            if value["fromCol"] == getColNumber(col) and value["fromRow"] == row-1:
                chartS = value
    if not chartS and flag == 1:
        for key,value in drawingsS[sheet].items():
            if "chart" in key: # get the last chart that is in the worksheet
                chartS = value
                
    if not chartS:
        print("There is no chart in the",sheet,"of the student's spreadsheet")
        return False
    if twoCharts: # twoCharts is either bar or line (it is not False)
        if not twoChartsFeature:
            if chartT["type"] == "twoCharts":
                if chartT.get(twoCharts,"empty") == "empty":
                    print("None of the charts in the teacaher's spreadsheet",sheet,"is a",twoCharts)
                    return np.nan
                else:
                    chartT = chartT[twoCharts]
            else:
                print("In the teacher's spreadsheet the chart does not contain two types")
                return np.nan
            if chartS["type"] == "twoCharts": # if not, let the oneChart be there
                if chartS.get(twoCharts,"empty") == "empty":
                    return False
                else:
                    chartS = chartS[twoCharts]
       
    return chartT,chartS

def checkAxes(sheet,col,row,twoCharts=False,twoChartsFeature=False):
    if isinstance(checkChart(sheet,col,row,twoCharts,twoChartsFeature),tuple):
        chartT,chartS = checkChart(sheet,col,row,twoCharts,twoChartsFeature)
        if chartT.get("axes","empty") == "empty":
            print("The chart in the teacher's spreadsheet don't have the feature axes")
            return np.nan
        if chartS.get("axes","empty") == chartT["axes"]:
            return True
    return False   
        
def checkChartStartCell(sheet,col,row,twoCharts=False,twoChartsFeature=False):
    if isinstance(checkChart(sheet,col,row,twoCharts,twoChartsFeature),tuple):
        chartT,chartS = checkChart(sheet,col,row,twoCharts,twoChartsFeature)
        if chartT["fromCol"] == chartS["fromCol"] and chartT["fromRow"] == chartS["fromRow"]:
            return True
    return False     
    
def checkChartTitle(sheet,col,row,twoCharts=False,twoChartsFeature=False):
    if isinstance(checkChart(sheet,col,row,twoCharts,twoChartsFeature),tuple):
        chartT,chartS = checkChart(sheet,col,row,twoCharts,twoChartsFeature)
        if checkData(chartT["title"],chartS["title"]):
            return True
    return False  
 
def checkChartColors(sheet,col,row,ND=False,twoCharts=False,twoChartsFeature=False):
    if isinstance(checkChart(sheet,col,row,twoCharts,twoChartsFeature),tuple):
        chartT,chartS = checkChart(sheet,col,row,twoCharts,twoChartsFeature)
        if chartT.get("type","empty") == "bar" or chartT.get("type","empty") == "col":
            if not ND:
                if chartT.get("color","empty") == chartS.get("color","empty"):
                    return True
            else:
                if chartT.get("color","empty") != chartS.get("color","empty"):
                    return True
    return False 

def checkChartType(sheet,col,row,twoCharts=False,twoChartsFeature=False):
    if isinstance(checkChart(sheet,col,row,twoCharts,twoChartsFeature),tuple):
        chartT,chartS = checkChart(sheet,col,row,twoCharts,twoChartsFeature)
        if chartT.get("type","empty") == chartS.get("type","empty"):
            return True
    return False 

def checkChartRef(sheet,col,row,twoCharts=False,twoChartsFeature=False):
    if isinstance(checkChart(sheet,col,row,twoCharts,twoChartsFeature),tuple):
        chartT,chartS = checkChart(sheet,col,row,twoCharts,twoChartsFeature)
        if chartT.get("type","empty") == "bar" or chartT.get("type","empty") == "col":
            if chartS.get("ref1","empty") != "empty" and chartS.get("ref2","empty") != "empty":
                ref1T = chartT["ref1"]
                ref1S = chartS["ref1"]
                res = []
                for r1T in ref1T:
                    if len(ref1S) == 1:
                        if ',' in ref1S[0]:
                            ref1S = ref1S[0].replace("(","").replace(")","").split(',')                            
                        
                    if any(r1T.split('!')[1] == r1S.split('!')[1] for r1S in ref1S):
                        res.append(True)
                if res == []:
                    res.append(False)
                ref2T = chartT["ref2"]
                ref2S = chartS["ref2"]
                for r2T in ref2T:
                    if len(ref2S) == 1:
                        if ',' in ref2S[0]:
                            ref2S = ref2S[0].replace("(","").replace(")","").split(',')  
                    if any(r2T.split('!')[1] == r2S.split('!')[1] for r2S in ref2S):
                        res.append(True)
                if len(res)==1:
                    res.append(False)
                return res
            else:
                print("At least one of the reference is not defined for the student's chart")
        if chartT.get("type","empty") == "line":
            if chartS.get("ref","empty") != "empty":
                refT = chartT["ref"]
                refS = chartS["ref"]
                res = []
                for rT in refT:
                    if any(rT.split('!')[1] == rS.split('!')[1] for rS in refS):
                        res.append(True)
                if res == []:
                    res.append(False)
                return res
            
    return [False]  


def checkPivotTable(area,startCol,startRow,element=None): # note that the element should be uniquely described
    startCol = startCol.upper()
    psT = getPivotTables(pivotTablesT,pivotCacheT)
    psS = getPivotTables(pivotTablesS,pivotCacheS)    
    if psT == {}:
        print("There is no pivot table in the teacher's spreadsheet")
        return [np.nan]
    pT = None
    for key,value in psT.items():
        if startCol+str(startRow)== key.split(':')[0]:
            pT = psT[key]
    if not pT: 
        print("There is no pivot table in the teacher's spreadsheet located on",startCol+str(startRow))
        return [np.nan]
    if pT.get(area,"empty") == "empty" and area != "inserted":
        print("The pivot table located on",startCol+str(startRow),"of the teacher's spreadsheet does not have",area)
        return [np.nan]
    if psS == {}:
        return [False]
    pS = None
    for key,value in psS.items():
        if startCol+str(startRow)== key.split(':')[0]:
            pS = psS[key]    
    if not pS:
        l= [1,0,-1]
        for i in l:
            for j in l:
                for key,value in psS.items():
                    if chr(ord(startCol)+i)+str(startRow+j)== key.split(':')[0]:
                        pS = psS[key] 
    if not pS:
        return [False]
    if area == "inserted":
        return [True]
    if pS.get(area,"empty") == "empty":
        return [False]
    if not element: 
        res = []
        for item in pT[area]:
            if any(checkData(item,itemS) for itemS in pS[area]):
                res.append(True)
            else:
                res.append(False)        
        return res
    else:
        element = element.lower()
        for item in pT[area]:
            if element in item:    
                if any(item in itemS for itemS in pS[area]):
                    return [True]
    return [False]

def getTable(worksheets,tables,sheet): # the given sheet should have only one table in it
    if tables != []:            # if the number of tables in the previous sheets are different it cannot return the desired table
        c = 0
        for w,worksheet in enumerate(worksheets):              
            if containsTable(worksheet) != 0:
                if w == int(sheet[-1])-1:
                    return tables[c]
                c += int(containsTable(worksheet))
                
def checkTableFilter(sheet,col): # returns true if all the filters on that col are the same
    col = col.upper()
    colNum = ord(col)-65
    tableT = getTable(worksheetsT,tableFilesT,sheet)
    tableS = getTable(worksheetsS,tableFilesS,sheet)
    if tableT:
        filtersT = tableFilters(tableT)
    else:
        print("There is no table in",sheet,"of the teacher's spreadsheet!")
        return np.nan
    if not tableS:
        return False
    filtersS = tableFilters(tableS)
    if filtersT.get(colNum) == "empty":
        print("There is no filter on column",col,"of",sheet,"of the teacher's spreadsheet")
        return np.nan
    if filtersS.get(colNum) == "empty":
        return False
    return filtersT[str(colNum)] == filtersS[str(colNum)]


# ### Preprocessing

# In[63]:


def identifyHiddenSheet(workbook):
    theWBook = accessWorkbook(workbook)
    sheets = theWBook.getElementsByTagName("sheets")[0].getElementsByTagName("sheet")
    res = []
    for i,sheet in enumerate(sheets):
        if sheet.hasAttribute("state"):
            if sheet.getAttribute("state") == "hidden":
                res.append(i)
    if res != []:
        return res
    return False


# ### Match Functions

# In[64]:


def getCell(row,colRef,rowRef):
    colRef = colRef.upper()
    if findCell(row[0][0])[1] == rowRef:
        for cell in row:
            if cell[0] == colRef + str(rowRef):
                return [cell[0],cell[1],cell[2],cell[3]]  

def getTCell(sheet,column,row):
    i = 1
    if len(sheetMatT[sheet]) > row:
        r = row-1
        rowT = sheetMatT[sheet][r]        
        if findCell(rowT[0][0])[1] != row:
            t = r-1
            while t>=0:
                rowT = sheetMatT[sheet][t] 
                if findCell(rowT[0][0])[1] != r:
                    t -= 1
                else:
                    break        
    else:
        rowT = sheetMatT[sheet][-1]
        r = len(sheetMatT[sheet])-1
    d = getCell(rowT,column,row)
    while not isinstance(d,list):
        rowT = sheetMatT[sheet][r-i]
        if r-i <= 0: # negative happens if we don't have the cell in an existing row
            return np.nan
        d = getCell(rowT,column,row)
        i += 1
    return d            
        
        
def getSCell(sheet,column,row):
    i = 1
    if len(sheetMatS[sheet]) > row:
        r = row-1
        rowS = sheetMatS[sheet][r]        
        if findCell(rowS[0][0])[1] != row:
            t = r-1
            while t>=0:
                rowS = sheetMatS[sheet][t] 
                if findCell(rowS[0][0])[1] != r:
                    t -= 1
                else:
                    break   
    else:
        if len(sheetMatS[sheet]) == 0:
            return np.nan
        rowS = sheetMatS[sheet][-1]
        r = len(sheetMatS[sheet])-1
    d = getCell(rowS,column,row) # a four-element cell
    while not isinstance(d,list):
        rowS = sheetMatS[sheet][r-i]
        if r-i <= 0:
            return np.nan
        d = getCell(rowS,column,row)
        i += 1
    return d 


def getTData(sheet,column,row):
    if isinstance(getTCell(sheet,column,row),list):
        return getTCell(sheet,column,row)[2]


def getSData(sheet,column,row):
    if isinstance(getSCell(sheet,column,row),list):
        return getSCell(sheet,column,row)[2]
    
            
def getSFormula(sheet,column,row):
    if isinstance(getSCell(sheet,column,row),list):
        return getSCell(sheet,column,row)[3]
    
def cellAddress(string):
    if findCell(string):
        if findCell(string)[0]+str(findCell(string)[1]) == string: # a simple cell address
            return True
    return False

def removeDollorSign(word):
    regex = re.compile("[$]")
    word = regex.sub('', word)
    return word 
        
def formulaInclude(formula,strings):  
    if formula:
        formula = formula.lower()
        for string in strings:
            string = string.lower()
            if string in formula:
                pass
            elif cellAddress(string): # flexible when the student used an absolute reference when being absolute was not needed
                if string in removeDollorSign(formula):
                    pass
                else:
                    return False
            else:
                return False
        return True
    return False


def formulaNotInclude(formula,strings):  
    if formula:
        formula = formula.lower()
        for string in strings:
            string = string.lower()
            if string in formula:
                return False         
        return True
    return False

    
def matchCell(sheet,column,row):
    column = column.upper()
    if len(sheetMatT)<int(sheet[-1]):
        print("There is no",sheet,"in the teacher's document")
        return np.nan
    if len(sheetMatS)<int(sheet[-1]):
        return False
    res = [] #TCell,SCell
    if isinstance(getTCell(sheet,column,row),list):
        res.append(getTCell(sheet,column,row))
    else:
#         print("Cell",column,row,"is not defined in the teacher's document",sheet)
        return np.nan
    if isinstance(getSCell(sheet,column,row),list):
        res.append(getSCell(sheet,column,row))        
    else:
        return False
    return res

        
def matchValue(sheet,column,row):
    res = matchCell(sheet,column,row)
    if res == False:
        return False
    if not isinstance(res,list):
        return np.nan # 0
    else:
        return checkData(res[0][2],res[1][2])
    
def matchValueE(sheet,column,row):
    res = matchCell(sheet,column,row)
    if res == False:
        return False
    if not isinstance(res,list):
        return np.nan # 0
    else:
        return res[0][2] == res[1][2]
    
def matchNValue(sheet,column,row):
    res = matchCell(sheet,column,row)
    if res == False:
        return False
    if not isinstance(res,list):
        return np.nan # 0
    else:
        if res[1][2] == None:
            return True
    return False
        

def matchStyle(sheet,column,row,func):
    column = column.upper()
    res = matchCell(sheet,column,row)
    if res == False:
        return [False]
    elif isinstance(res,list):
        return func(res[0][1],res[1][1])
    return [np.nan]
    
    
def matchFormula(sheet,column,row,E=False):
    res = matchCell(sheet,column,row)
    if res == False:
        return False
    if not isinstance(res,list):
        return np.nan
    elif E == False:
        return checkFormula(res[0][3],res[1][3])
    else:
        return checkExactFormula(res[0][3],res[1][3])
    
    
def styleListRes(sheet,column,row,func,ind):
    resList = matchStyle(sheet,column,row,func)
    if len(resList) == 1:
        if resList == [True]:
            return True
        elif resList == [False]:
            return False
        elif resList == [np.nan]:
            return np.nan
        else:
            print("This case should not happen!")
    else:
        return resList[ind]          


# ### Grade Map

# In[65]:


def calGrade(match):
    full = []
    
    # to calculate the full grade
    for item in match:
        if item == False:
            full.append(True)
        else:
            full.append(item)

    if match == [None] or np.nansum(full) == 0:
        return np.nan    
    elif match == full:
        return "full-mark"
    else:
        res = np.around(np.nansum(match)/np.nansum(full),2)
        return math.trunc(res*100)/100

def calBGrade(match):
    if match == True:
        return "full-mark"
    elif match == None or np.isnan(match):
        return np.nan
    else:
        return 0

def chartProcess(sheet):
    if chartsT.get(sheet,"empty") == "empty":
        print("You probably has entered the sheet incorrectly.")
        return np.nan
    else:
        if chartsS.get(sheet,"empty") == "empty":
            return False
        else:
            return True


# In[66]:


def grade(*args):
    if len(args) == 1:
        if args[0] == "theme":
            return calBGrade(checkThemes(themeNameT,themeNameS))
        
        elif args[0] == "sheetNames":
            return calGrade(checkSheetNames(sheetNamesT,sheetNamesS))
        
        else:
            print("Invalid usage of one-argument function! The Argument should be any of: theme, sheetNames (for all sheets)")
            
    elif len(args) == 2:  
        if args[0] == "defName":
            return calGrade(checkDefinedName(args[1]))
            
        if args[1] == "sheetName":
            ind = int(args[0][-1])-1
            res = checkSheetNames(sheetNamesT,sheetNamesS)            
            return calBGrade(res[ind])
        
        elif args[1] == "noHLink":
            return calBGrade(checkNoHyperlink(args[0]))
        
        elif args[1] == "colW":         
            return calGrade(checkColsW(args[0]))
        
        elif args[1] == "colHide":         
            return calGrade(checkColsHide(args[0]))
        
        elif args[1] == "freeze":         
            return calBGrade(checkFreeze(args[0]))
        
        elif args[1] == "sheetProtect":         
            return calBGrade(checkSheetProtect(args[0]))
        
        elif args[1] == "mergedCell":
            return calBGrade(checkMergeCells(args[0]))
        
        elif args[1] == "table": # if there exists a table in the specified sheet
            return calBGrade(checkWorksheetTable(args[0]))
        
        elif args[1] == "tableRef": # if the table reference is correct
            return calBGrade(checkTableRef(args[0]))
            
        elif args[1] == "orientation":
            return calBGrade(checkOrientation(args[0]))
        
        elif args[1] == "printPage":
            return calBGrade(checkPrintPage(args[0]))
        
        
        else:
            print("Invalid usage of two-argument function! The first Argument is the sheet (e.g. 'sheet1')            The second argument should be any of these: sheetName (for all sheets), colW (checking width of all columns),            colHide (checking the hidden columns of a sheet), freeze, sheetProtect, chart (to check chartType, group and             references), chartType, chartSize, chartTitle, chartStartCell, chartColors, chartColorNotDefault, orientation")
            
    elif len(args) == 3:       
        if args[1] == "colW":         
            return calGrade(checkColsW(args[0],args[2]))
        
        elif args[1] == "tableFilter":
            return calBGrade(checkTableFilter(args[0],args[2]))
        
        elif args[1] == "conFormat":         
            return calGrade(checkDXFs(args[0],args[2]))
        
        elif args[1] == "conFormatT":         
            return calBGrade(checkDXFs(args[0],args[2])[0])
        
#         elif args[1] == "conFormatF":         
#             return calBGrade(checkDXFs(args[0],args[2])[1])
        
        elif args[1] == "rowHeight":         
            return calBGrade(checkRowHeight(args[0],args[2]))  
        
        elif args[1] == "picture":  
            # other features can be added
            if args[2] == "weight" or args[2] == "height" or args[2] == "startCell" or args[2] == "size" or args[2] == "imgEffect"            or args[2] == "recolor" or args[2] == "exist" or args[2] == "inserted": # exist has more criteria (within 3 cells)
                return calBGrade(checkPicture(args[0],args[2]))   
            
        elif args[1] == "header" or args[1] == "footer":
            if args[2] == "left" or args[2] == "right" or args[2] == "center":
                return calBGrade(checkHeaderFooter(args[0],args[1],args[2]))
        
        else:
            print("Invalid usage of three-argument function! The first Argument is the sheet (e.g. sheet1)            The second and third arguments should be any of these: (colW,int) column width and the number of columns that             needs to be checked, (conFormat,column), (conFormatT,column) checking only type in case the point was different than            formating, (conFormatF,column) formatting            ,             ")


    # check style on a cell
    elif len(args) == 4:
        if args[0] == "pivot":
            if args[1] == "rows" or args[1] == "columns" or args[1] == "values" or args[1] == "filters" or args[1] == "inserted":
                return calGrade(checkPivotTable(args[1],args[2],args[3]))
            
        else:
            # arguments: sheet, keyword, column, row        
            if args[1] == "value":
                return calBGrade(matchValue(args[0],args[2],args[3]))
            
            elif args[1] == "valueN": # value is none
                return calBGrade(matchNValue(args[0],args[2],args[3]))
            
            elif args[1] == "valueE": # the exact value
                return calBGrade(matchValueE(args[0],args[2],args[3]))            

            elif args[1] == "format":
                return calBGrade(matchStyle(args[0],args[2],args[3],checkNumFmt))

            elif args[1] == "protect":
                return calBGrade(matchStyle(args[0],args[2],args[3],checkCellProtection))

            elif args[1] == "formulaE": # formula Exact
                return calBGrade(matchFormula(args[0],args[2],args[3],True))    

            elif args[1] == "formulaF": # formula Flexible. only works on one cell
                if matchFormula(args[0],args[2],args[3]) == True: # formula is correct but the answer isn't because of previous questions
                    return calBGrade(True)
                elif matchFormula(args[0],args[2],args[3]) != False: # a formula is used
                    return calBGrade(matchValue(args[0],args[2],args[3]))
                else:
                    return 0 

            elif args[1] == "align":
                return calGrade(matchStyle(args[0],args[2],args[3],checkAlign))

            # style for left, right, top, bottom, diagonal
            elif args[1] == "border":
                return calBGrade(matchStyle(args[0],args[2],args[3],checkBorder))

            elif args[1] == "fill":
                return calGrade(matchStyle(args[0],args[2],args[3],checkFill))

            # 0.bold, 1.underline, 2.italic, 3.strike, 4.size, 5.colorT, 6.colorRGB, 7.name, 8.family, 9.scheme
            elif args[1] == "font":
                return calGrade(matchStyle(args[0],args[2],args[3],checkFont))

            elif args[1] == "bold":
                return calBGrade(styleListRes(args[0],args[2],args[3],checkFont,0))

            elif args[1] == "bold&size":
                res = [calBGrade(styleListRes(args[0],args[2],args[3],checkFont,0)),calBGrade(styleListRes(args[0],args[2],args[3],checkFont,4))]
                return res

            elif args[1] == "underline":
                return calBGrade(styleListRes(args[0],args[2],args[3],checkFont,1))

            elif args[1] == "italic":
                return calBGrade(styleListRes(args[0],args[2],args[3],checkFont,2))

            elif args[1] == "size":
                return calBGrade(styleListRes(args[0],args[2],args[3],checkFont,4))

            elif args[1] == "color":
                return calBGrade(styleListRes(args[0],args[2],args[3],checkFont,6))

            elif args[1] == "fontName":
                return calBGrade(styleListRes(args[0],args[2],args[3],checkFont,7))

            elif args[1] == "family":
                return calBGrade(styleListRes(args[0],args[2],args[3],checkFont,8))

            elif args[1] == "dataValidation":
                return calBGrade(checkDataValidation(args[0],args[2],args[3]))

            elif args[1] == "chartType":
                return calBGrade(checkChartType(args[0],args[2],args[3]))

            elif args[1] == "chartRef": # partial point for each ref
                return calGrade(checkChartRef(args[0],args[2],args[3]))

            elif args[1] == "chartColor":
                return calBGrade(checkChartColors(args[0],args[2],args[3]))

            elif args[1] == "chartColorND":
                return calBGrade(checkChartColors(args[0],args[2],args[3],True))

            elif args[1] == "chartTitle":
                return calBGrade(checkChartTitle(args[0],args[2],args[3]))

            elif args[1] == "chartStartCell":
                return calBGrade(checkChartStartCell(args[0],args[2],args[3]))    

            elif args[1] == "picture":  
                if args[2] == "weight" or args[2] == "height" or args[2] == "startCell" or args[2] == "size" or args[2] == "imgEffect"                or args[2] == "recolor" or args[2] == "exist" or args[2] == "inserted":
                    return calBGrade(checkPicture(args[0],args[2],args[3])) 

            else:
                print("Invalid Arguments! please enter the sheet (e.g. 'sheet1') following by any of these as the second argument:                        value, format (numberFormat), protect (cell protection), formulaE (exact formula check),                         formulaF (flexible formula check), align, border (border style), fill,                       font (to check the whole font attributes), bold, underline, italic, size (font size),                       color (font RGB color), name (font name), family (font family), dv (data validation) following by a cell                        i.e. column like 'a' and row like 1 (grade('sheet1','size','a',1))")
            
            
    elif len(args) == 5:
        if args[0] == "pivot": # args[4] should be uniquely identified
            if args[1] == "rows" or args[1] == "columns" or args[1] == "values" or args[1] == "filters" or args[1] == "inserted":
                return calGrade(checkPivotTable(args[1],args[2],args[3],args[4]))
        else:
            if args[1] == "formulaF":
                if matchFormula(args[0],args[2],args[3]) == True: # formula is exactly the same
                    return calBGrade(True)
                elif matchFormula(args[0],args[2],args[3]) != False: # a formula is used
                    if formulaInclude(getSFormula(args[0],args[2],args[3]),args[4]):
                        return calBGrade(matchValue(args[0],args[2],args[3]))
                return 0 

            elif args[1] == "formulaFNot":
                if matchFormula(args[0],args[2],args[3]) == True: # formula is exactly the same
                    return calBGrade(True)
                elif matchFormula(args[0],args[2],args[3]) != False: # a formula is used
                    if formulaNotInclude(getSFormula(args[0],args[2],args[3]),args[4]):
                        return calBGrade(matchValue(args[0],args[2],args[3]))
                return 0 
            
            # checking in twoCharts
            elif args[1] == "chartType":
                return calBGrade(checkChartType(args[0],args[2],args[3],args[4]))
            
            elif args[1] == "chartRef": # partial point for each ref
                return calGrade(checkChartRef(args[0],args[2],args[3],args[4]))

            elif args[1] == "chartColor":
                return calBGrade(checkChartColors(args[0],args[2],args[3],args[4]))

            elif args[1] == "chartColorND":
                return calBGrade(checkChartColors(args[0],args[2],args[3],True,args[4]))

            elif args[1] == "chartTitle": #title is a common feature of a combined chart
                return calBGrade(checkChartTitle(args[0],args[2],args[3],args[4],True))
            
            elif args[1] == "chartAxes":
                return calBGrade(checkAxes(args[0],args[2],args[3],args[4],True))

            elif args[1] == "chartStartCell":
                return calBGrade(checkChartStartCell(args[0],args[2],args[3],args[4],True))
            else:
                print("Invalid Arguments!")
            
            
    elif len(args) == 6:
        if args[1] == "formulaF": # check if a formula is used and the value is entirely, 100%, correct
            for i in range(args[5]-args[3]+1): # iterating on rows
                for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                    if matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) == True: 
                        pass
                    elif matchFormula(args[0],args[2],args[3]) != False: # a formula is used
                        temp = calBGrade(matchValue(args[0],chr(ord(args[2].upper())+j),args[3]+i))
                        if temp == False:
#                             print(chr(ord(args[2].upper())+j),args[3]+i)
                            return 0
                    else:
                        return 0             
            return "full-mark"
        
        elif args[1] == "value": # check if the value is 100% correct
            for i in range(args[5]-args[3]+1): # iterating on rows
                for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                    temp = calBGrade(matchValue(args[0],chr(ord(args[2].upper())+j),args[3]+i))
                    if temp == False:
#                         print(chr(ord(args[2].upper())+j),args[3]+i)
                        return 0
            return "full-mark"
        
        elif args[1] == "format": # check if the number format is 100% correct
            for i in range(args[5]-args[3]+1): # iterating on rows
                for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                    temp = grade(args[0],"format",chr(ord(args[2].upper())+j),args[3]+i)
                    if temp == False:
                        return 0
            return "full-mark"
            
        elif args[1] == "formulaFPartial": # returns partial credit for the formulaF on the specified range
            res = []
            if args[2].lower() == args[4].lower(): # a range on a column
                for i in range(abs(args[3]-args[5])+1):
                    res.append(grade(args[0],"formulaF",args[2],args[3]+i))
                return res
            if args[3] == args[5]: # a range on a row
                for i in range(abs(ord(args[2].lower())-ord(args[4].lower()))+1):
                    res.append(grade(args[0],"formulaF",chr(ord(args[2].lower())+i),args[3]))
                return res
            elif (args[2].lower() < args[4].lower() and args[3] < args[5]):
                for i in range(args[5]-args[3]+1):
                    temp = []
                    for j in range(ord(args[4].lower())-ord(args[2].lower())+1):
                        temp.append(grade(args[0],"formulaF",chr(ord(args[2].lower())+j),args[3]+i))
                    res.append(temp)
                return res            
            else:
                print("This is not a valid range. To get a rectangle start from the top left cell.")  
                
        elif args[1] == "valuePartial": # returns partial credit for the value on the specified range
            res = []
            if args[2].lower() == args[4].lower(): # a range on a column
                for i in range(abs(args[3]-args[5])+1):
                    res.append(grade(args[0],"value",args[2],args[3]+i))
                return res
            if args[3] == args[5]: # a range on a row
                for i in range(abs(ord(args[2].lower())-ord(args[4].lower()))+1):
                    res.append(grade(args[0],"value",chr(ord(args[2].lower())+i),args[3]))
                return res
            elif (args[2].lower() < args[4].lower() and args[3] < args[5]):
                for i in range(args[5]-args[3]+1):
                    temp = []
                    for j in range(ord(args[4].lower())-ord(args[2].lower())+1):
                        temp.append(grade(args[0],"value",chr(ord(args[2].lower())+j),args[3]+i))
                    res.append(temp)
                return res            
            else:
                print("This is not a valid range. To get a rectangle start from the top left cell.")  
                
        else:
            res = []
            if args[2].lower() == args[4].lower(): # a range on a column
                for i in range(abs(args[3]-args[5])+1):
                    res.append(grade(args[0],args[1],args[2],args[3]+i))
                return res
            if args[3] == args[5]: # a range on a row
                for i in range(abs(ord(args[2].lower())-ord(args[4].lower()))+1):
                    res.append(grade(args[0],args[1],chr(ord(args[2].lower())+i),args[3]))
                return res
            elif (args[2].lower() < args[4].lower() and args[3] < args[5]):
                for i in range(args[5]-args[3]+1):
                    temp = []
                    for j in range(ord(args[4].lower())-ord(args[2].lower())+1):
                        temp.append(grade(args[0],args[1],chr(ord(args[2].lower())+j),args[3]+i))
                    res.append(temp)
                return res            
            else:
                print("This is not a valid range. To get a rectangle start from the top left cell.")  
                
    elif len(args) == 7:
        if args[1] == "value": # check if the value is args[6] correct
            c1 = 0 # count all cells
            c2 = 0 # count cells that don't match
            for i in range(args[5]-args[3]+1): # iterating on rows
                for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                    temp = calBGrade(matchValue(args[0],chr(ord(args[2].upper())+j),args[3]+i))
                    c1 += 1
                    if temp == False:
#                         print(chr(ord(args[2].upper())+j),args[3]+i)
                        c2 +=1
            if c2/c1 < 1-args[6]:
                return "full-mark"
            return 0
        
        elif args[1] == "formulaF": # check if a formula is used and the value is args[6] correct or if formula includes the list in args[6]
            if isinstance(args[6],list):
                for i in range(args[5]-args[3]+1): # iterating on rows
                    for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                        if matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) == True:
                            pass
                        elif matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) != False:# a formula is used and has args[6]
                            if formulaInclude(getSFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i),args[6]):                         
                                temp = calBGrade(matchValue(args[0],chr(ord(args[2].upper())+j),args[3]+i))
                                if temp == False:
#                                     print(chr(ord(args[2].upper())+j),args[3]+i)
                                    return 0    
                            else:
                                return 0
                        else:
                            return 0  
                return "full-mark"
            
            elif isinstance(args[6],float):
                c1 = 0
                c2 = 0
                for i in range(args[5]-args[3]+1): # iterating on rows
                    for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                        c1 += 1
                        if matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) == True: 
                            pass
                        elif matchFormula(args[0],args[2],args[3]) != False: # a formula is used
                            temp = calBGrade(matchValue(args[0],chr(ord(args[2].upper())+j),args[3]+i))
                            if temp == False:
    #                             print(chr(ord(args[2].upper())+j),args[3]+i)
                                c2 +=1
                        else:
                            c2 +=1 
                    if c2/c1 > 1-args[6]:# to check after each row
                        return 0
                if c2/c1 > 1-args[6]:
                    return 0
                return "full-mark"
        
        elif args[1] == "format": # check if the number format is args[6] correct
            c1 = 0
            c2 = 0
            for i in range(args[5]-args[3]+1): # iterating on rows
                for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                    c1 += 1
                    temp = grade(args[0],"format",chr(ord(args[2].upper())+j),args[3]+i)
                    if temp == False:
                        c2 += 1
                if c2/c1 > 1-args[6]: # to check after each row
                    return 0
            if c2/c1 > 1-args[6]:
                return 0
            return "full-mark"

        
        elif args[1] == "formulaFNot":# a formula is used that does not have any of elements of args[6] 
            for i in range(args[5]-args[3]+1): # iterating on rows
                for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                    if matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) == True:
                        pass
                    elif matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) != False:# a formula is used and has args[6]
                        if formulaNotInclude(getSFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i),args[6]):                         
                            temp = calBGrade(matchValue(args[0],chr(ord(args[2].upper())+j),args[3]+i))
                            if temp == False:
                                return 0 
                        else:
                            return 0
                    else:
                        return 0  
            return "full-mark"    
            
        elif args[1] == "formulaFNotPartial": # returns partial credit for the formulaFNOt on the specified range
            res = []
            if args[2].lower() == args[4].lower(): # a range on a column
                for i in range(abs(args[3]-args[5])+1):
                    res.append(grade(args[0],"formulaFNot",args[2],args[3]+i,args[6]))
                return res
            if args[3] == args[5]: # a range on a row
                for i in range(abs(ord(args[2].lower())-ord(args[4].lower()))+1):
                    res.append(grade(args[0],"formulaFNot",chr(ord(args[2].lower())+i),args[3],args[6]))
                return res
            elif (args[2].lower() < args[4].lower() and args[3] < args[5]):
                for i in range(args[5]-args[3]+1):
                    temp = []
                    for j in range(ord(args[4].lower())-ord(args[2].lower())+1):
                        temp.append(grade(args[0],"formulaFNot",chr(ord(args[2].lower())+j),args[3]+i,args[6]))
                    res.append(temp)
                return res            
            else:
                print("This is not a valid range. To get a rectangle start from the top left cell.")
                
        elif args[1] == "formulaFPartial": # returns partial credit for the formulaF on the specified range having args[6]
            res = []
            if args[2].lower() == args[4].lower(): # a range on a column
                for i in range(abs(args[3]-args[5])+1):
                    res.append(grade(args[0],"formulaF",args[2],args[3]+i,args[6]))
                return res
            if args[3] == args[5]: # a range on a row
                for i in range(abs(ord(args[2].lower())-ord(args[4].lower()))+1):
                    res.append(grade(args[0],"formulaF",chr(ord(args[2].lower())+i),args[3],args[6]))
                return res
            elif (args[2].lower() < args[4].lower() and args[3] < args[5]):
                for i in range(args[5]-args[3]+1):
                    temp = []
                    for j in range(ord(args[4].lower())-ord(args[2].lower())+1):
                        temp.append(grade(args[0],"formulaF",chr(ord(args[2].lower())+j),args[3]+i,args[6]))
                    res.append(temp)
                return res            
            else:
                print("This is not a valid range. To get a rectangle start from the top left cell.")
                
        else:
            print("Invalid 7-Argument use of the grade function.")
    
    elif len(args) == 8:
        c1 = 0
        c2 = 0
        if args[1] == "formulaF":# a formula is used that has args[6] and the values are args[7] correct
            for i in range(args[5]-args[3]+1): # iterating on rows
                for j in range(ord(args[4].lower())-ord(args[2].lower())+1): # iterating on columns
                    c1 += 1
                    if matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) == True:
                        pass
                    elif matchFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i) != False:# a formula is used and has args[6]
                        if formulaInclude(getSFormula(args[0],chr(ord(args[2].upper())+j),args[3]+i),args[6]):                         
                            temp = calBGrade(matchValue(args[0],chr(ord(args[2].upper())+j),args[3]+i))
                            if temp == False:
#                                 print(chr(ord(args[2].upper())+j),args[3]+i)
                                c2 +=1   
                        else:
                            c2 +=1 
                    else:
                        c2 +=1  
                    if c2/c1 > 1-args[7]:# to check after each row
                        return 0
            if c2/c1 > 1-args[7]:
                    return 0
            return "full-mark"
        
        else:
            print("Invalid 8-Argument use of the grade function.")
            
    else:
        print("Wrong number of inputs. Enter the sheet No. followed by a keyword argument")



# In[67]:

def finalExam(filename):
    ans = []
    qLabel = []
    point = []
    
    # how to check question 1 or 2 :D
    ans.append("full-mark")
    qLabel.append("1")
    point.append(1)
    
    ans.append("full-mark")
    qLabel.append("2")
    point.append(1)

    ans.append(grade("theme"))
    qLabel.append("3")
    point.append(1)

    ans.append(grade("sheet1","bold&size","a",1,"k",1))
    qLabel.append("4.a")
    point.append(2)

    ans.append(grade("sheet1","colW",9)) # other columns are not created yet
    qLabel.append("4.b")
    point.append(1)

    ans.append(grade("sheet1","colHide"))
    qLabel.append("4.c")
    point.append(2)

    ans.append(grade("sheet1","sheetName"))
    qLabel.append("4.d")
    point.append(1)

    ans.append(grade("sheet1","sheetProtect"))
    qLabel.append("4.e")
    point.append(1)

    ans.append(grade("sheet2","sheetName"))
    qLabel.append("5.a")
    point.append(1)

    ans.append(grade("sheet2","freeze"))
    qLabel.append("5.b")
    point.append(1)

    ans.append(grade("sheet2","conFormat","H")) # if the grades are the same it can be written in one command
    qLabel.append("6")
    point.append(1)

    ans.append(grade("sheet2","conFormat","G")) 
    qLabel.append("7")
    point.append(1)

    ans.append(grade("sheet2","value","a",6261,"a",6264)) 
    qLabel.append("8")
    point.append(1)
    
    ans.append(grade("sheet2","bold&size","a",6261,"a",6264)) 
    qLabel.append("8.a")
    point.append(2)

    ans.append(grade("sheet2","formulaE","b",6261,"f",6264)) 
    qLabel.append("8.b")
    point.append(4)

    ans.append(grade("sheet2","border","a",6261,"f",6264)) 
    qLabel.append("8.c")
    point.append(1)

    ans.append(grade("sheet2","format","c",6262,"f",6262)) 
    qLabel.append("8.d")
    point.append(1)

    ans.append(grade("sheet2","value","j",1)) 
    qLabel.append("9.a")
    point.append(1)

    ans.append(grade("sheet2","bold&size","j",1)) 
    qLabel.append("9.b")
    point.append(1)

    ans.append(grade("sheet2","formulaF","j",2,"j",6260)) 
    qLabel.append("9.c")
    point.append(1)

    ans.append(grade("sheet2","format","j",2,"j",6260))
    qLabel.append("9.d")
    point.append(1)

    ans.append(grade("sheet2","value","k",1)) 
    qLabel.append("10.a")
    point.append(1)

    ans.append(grade("sheet2","formulaFValue","k",2,"k",6260,"expensive")) 
    qLabel.append("10.b")
    point.append(2)
    
    ans.append(grade("sheet2","formulaFValue","k",2,"k",6260,"inexpensive")) 
    qLabel.append("10.c")
    point.append(2)
    
    ans.append(grade("sheet2","formulaFValue","k",2,"k",6260,"affordable")) 
    qLabel.append("10.d")
    point.append(2)

    ans.append(grade("sheet2","value","l",1)) 
    qLabel.append("11.a")
    point.append(1)

    ans.append(grade("sheet2","formulaF","l",2,"l",6260)) 
    qLabel.append("11.b")
    point.append(1)

    ans.append(grade("sheet2","format","l",2,"l",6260)) 
    qLabel.append("11.c")
    point.append(1)

    ans.append(grade("sheet3","sheetName"))
    qLabel.append("12.a")
    point.append(1)

    ans.append(grade("sheet3","formulaValue","l",2,"l",6260))
    qLabel.append("12.b")
    point.append(2)

    ans.append(grade("sheet3","conFormat","D")) 
    qLabel.append("12.c")
    point.append(1)

    ans.append(grade("sheet3","value","n",1))
    qLabel.append("13.a")
    point.append(1)

    ans.append(grade("sheet3","bold&size","n",1))
    qLabel.append("13.b")
    point.append(1)

    ans.append(grade("sheet3","formulaF","n",2,"n",6260)) 
    qLabel.append("13.c")
    point.append(2)
    
    resTable = printResult(qLabel,point,ans,filename)
    return resTable


# ### Printing the Results

# In[75]:


def dfResult(qLabel,point,res,filename,df):
    if len(df) == 0:
        df = pd.DataFrame(columns = np.hstack(('0',qLabel)))          
        df.loc[0] = np.zeros(len(qLabel)+1) # empty for description
        df.loc[1] = np.hstack(("Points",point)) 
        df.loc[2] = np.hstack((filename[0:7],res)) # the first student
    else:
        ind = len(df.index)
        df.loc[ind] = np.hstack((filename[0:7],res))
    return df


# In[76]:


def printResult(qLabel,point,ans,filename,df=[],desc=None):
    header = ['Question', 'Question Points', 'Points you have earned']    
    header2 = ['Question number','Question Description','Question Points', 'Points you have earned']
    jsonList = []
    with open('%s.csv' %filename, 'w', newline='') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=',',quotechar='|', quoting=csv.QUOTE_MINIMAL)
        if desc:
            spamwriter.writerow(header2)
        else:
            spamwriter.writerow(header)
        resTable = PrettyTable(['Question', 'Question Points', 'Points you have earned'])
        total = []
        for i in range(len(qLabel)):
            if isinstance(ans[i], list):
                full = []
                for counter in range(len(ans[i])):
                    if isinstance(ans[i][counter], list):
                        temp = []
                        for counter2 in range(len(ans[i][counter])):
                            if ans[i][counter][counter2] == "full-mark":
                                ans[i][counter][counter2] = 1
                            if np.isnan(ans[i][counter][counter2]):
                                temp.append(np.nan)
                            else:
                                temp.append(1)
                        full.append(np.nansum(temp))

                    else:
                        if ans[i][counter] == "full-mark":
                            ans[i][counter] = 1
                        if ans[i][counter] == 0:
                            full.append(1)
                        else:
                            full.append(ans[i][counter])
                if np.nansum(full) == 0:
                    total.append(0)
                else:
                    res = np.around((np.nansum(ans[i])/np.nansum(full))*point[i],2)
                    total.append(math.trunc(res*100)/100)
            else:
#                 print(qLabel[i],ans[i])
                if ans[i] == "full-mark":
                    total.append(point[i])
                elif np.isnan(ans[i]):
                    print("check the question: ",qLabel[i])
                    total.append(0)
                else:
                    res = point[i]*ans[i]
                    total.append(math.trunc(res*100)/100)
            theRow = [qLabel[i],point[i],total[-1]]
            if point[i] == 0:
                p = "Not Evaluated"
                e = "Not Evaluated"
            else:
                p = str(point[i])
                e = str(total[-1])
            
            jsonList.append({"question":qLabel[i],"points":p,"earned":e})
            if desc:
                theRow2 = [qLabel[i],desc[i],point[i],total[-1]]
                spamwriter.writerow(theRow2)
            else:
                spamwriter.writerow(theRow)
            resTable.add_row(theRow)
        df = dfResult(qLabel,point,total,filename,df)
        resTable.add_row(["----------","-----------------","------------------------"])  
        lastRow = ["Total",np.rint(np.sum(point)),np.around(np.sum(total),decimals=2)]
        lastRow2 = ["Total",'',np.rint(np.sum(point)),np.around(np.sum(total),decimals=2)]
        with open('%s.json' %filename, 'w') as outfile:
            jdata = {"origfile":filename+".xlsx","totalscore":str(100*np.sum(point)/np.sum(total)),"results":jsonList}
            json.dump(jdata, outfile, ensure_ascii=False,indent=4)
        if desc:
            spamwriter.writerow(lastRow2)
        else:
            spamwriter.writerow(lastRow)
        resTable.add_row(lastRow)        
        return resTable,df


# # Main

# In[77]:


# Setting up the teacher's document
# this is to remove the content of folder teacher, so it will be empty to be extracted to
if os.path.exists("teacher"):
    for the_file in os.listdir("teacher"):
        file_path = os.path.join("teacher", the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path): shutil.rmtree(file_path)
        except Exception as e:
            print(e)
        
# extracting to the folder named teacher        
xlsT = zipfile.ZipFile("TEACHER.xlsx", 'r')
xlsT.extractall("teacher")
xlsT.close()

os.chdir("teacher")
os.chdir("xl")
#calT = minidom.parse('calcChain.xml')
strT = None
if os.path.isfile('sharedStrings.xml'):
    strT = minidom.parse('sharedStrings.xml')
stylesT = minidom.parse('styles.xml')
workbookT = minidom.parse('workbook.xml')
worksheetsT = []
themesT = None
chartFilesT = []
drawingFilesT = []
tableFilesT = []
pivotCacheT = []
pivotTablesT = []

if os.path.exists("charts"):
    for file in os.listdir("charts"):
        if file.endswith(".xml"):
            if file.startswith("chart"):
                os.chdir("charts") # change directory to charts
                chartFilesT.append(minidom.parse(file))
                get_ipython().run_line_magic('cd', '..')

if os.path.exists("drawings"):
    for file in os.listdir("drawings"):
        if file.endswith(".xml"):
            if file.startswith("drawing"):
                os.chdir("drawings") # change directory to drawings
                drawingFilesT.append(minidom.parse(file))
                get_ipython().run_line_magic('cd', '..')

if os.path.exists("pivotCache"):
    for file in os.listdir("pivotCache"):
        if file.endswith(".xml"):
            if file.startswith("pivotCacheDefinition"):
                os.chdir("pivotCache") # change directory to pivotCache
                pivotCacheT.append(minidom.parse(file))
                get_ipython().run_line_magic('cd', '..')
                
if os.path.exists("pivotTables"):
    for file in os.listdir("pivotTables"):
        if file.endswith(".xml"):
            if file.startswith("pivotTable"):
                os.chdir("pivotTables") 
                pivotTablesT.append(minidom.parse(file))
                get_ipython().run_line_magic('cd', '..')
                
if os.path.exists("tables"):
    for file in os.listdir("tables"):
        if file.endswith(".xml"):
            if file.startswith("table"):
                os.chdir("tables") 
                tableFilesT.append(minidom.parse(file))
                get_ipython().run_line_magic('cd', '..')
                
for theme in os.listdir("theme"):
    if theme.endswith(".xml"):
        os.chdir("theme") # change directory to theme
        themesT = minidom.parse(theme)
        get_ipython().run_line_magic('cd', '..')
        
for sheet in os.listdir("worksheets"):
    if sheet.endswith(".xml"):
        os.chdir("worksheets") # change directory to worksheets
        worksheetsT.append(minidom.parse(sheet))
        get_ipython().run_line_magic('cd', '..')
get_ipython().run_line_magic('cd', '..')
get_ipython().run_line_magic('cd', '..')

stDicT = styleDictionary(stylesT)
dxfDicT = dxfsDictionary(stylesT)
strArrT = None
if strT:
    strArrT = sharedString(strT)
themeNameT = getTheme(themesT)
sheetNamesT = sheetNames(workbookT)
drawingsT = drawingFile(drawingFilesT,worksheetsT,chartFilesT)
# chartsT = assignChart(chartFilesT,drawingsT)
# dimensionT = worksheetDims(worksheetsT)
sheetMatT = worksheetMats(worksheetsT,strArrT)
colsWT = worksheetColsWRef(worksheetsT)
hidColsT = worksheetColHidden(worksheetsT)
freezeT = worksheetFreeze(worksheetsT)
conditionalFT = worksheetConditionalFormatting(worksheetsT)
sheetProtectT = worksheetProtection(worksheetsT)
# chartDicT = chartInSheet(chartsT,sheetNamesT)


# In[79]:


df = []
for studentFile in os.listdir("STUDENTFILES"):
    print(studentFile)
    os.chdir("STUDENTFILES")
    # Making the student folder empty
    if os.path.exists("student"):
        for the_file in os.listdir("student"):
            file_path = os.path.join("student", the_file)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception as e:
                print(e)
                
    # Extracting XML 
    if studentFile.endswith(".xlsx") and not studentFile.lstrip().startswith('~'):
        xlsS = zipfile.ZipFile(studentFile, 'r')
        xlsS.extractall("student")
        xlsS.close()

        # Parsing XML
        os.chdir("student")
        os.chdir("xl")
        #calS = minidom.parse('calcChain.xml')
        strS = None
        if os.path.isfile('sharedStrings.xml'):
            strS = minidom.parse('sharedStrings.xml')
        stylesS = minidom.parse('styles.xml')
        workbookS = minidom.parse('workbook.xml')
        worksheetsS = []
        themesS = None
        chartFilesS = []
        # chartColorsS = []
        # chartStylesS = []
        drawingFilesS = []
        tableFilesS = []
        pivotCacheS = []
        pivotTablesS = []

        for sheet in os.listdir("worksheets"):
            if sheet.endswith(".xml"):
                os.chdir("worksheets") # change directory to worksheets        
                worksheetsS.append(minidom.parse(sheet))
                get_ipython().run_line_magic('cd', '..')

        for theme in os.listdir("theme"):
            if theme.endswith(".xml"):
                os.chdir("theme") 
                themesS = minidom.parse(theme)
                get_ipython().run_line_magic('cd', '..')

        if os.path.exists("charts"):
            for file in os.listdir("charts"):
                if file.endswith(".xml"):
                    if file.startswith("chart"):
                        os.chdir("charts") 
                        chartFilesS.append(minidom.parse(file))
                        get_ipython().run_line_magic('cd', '..')
                        
        if os.path.exists("pivotCache"):
            for file in os.listdir("pivotCache"):
                if file.endswith(".xml"):
                    if file.startswith("pivotCacheDefinition"):
                        os.chdir("pivotCache") # change directory to pivotCache
                        pivotCacheS.append(minidom.parse(file))
                        get_ipython().run_line_magic('cd', '..')

        if os.path.exists("pivotTables"):
            for file in os.listdir("pivotTables"):
                if file.endswith(".xml"):
                    if file.startswith("pivotTable"):
                        os.chdir("pivotTables") 
                        pivotTablesS.append(minidom.parse(file))
                        get_ipython().run_line_magic('cd', '..')
                        
        if os.path.exists("tables"):
            for file in os.listdir("tables"):
                if file.endswith(".xml"):
                    if file.startswith("table"):
                        os.chdir("tables") 
                        tableFilesS.append(minidom.parse(file))
                        get_ipython().run_line_magic('cd', '..')

        if os.path.exists("drawings"):
            for file in os.listdir("drawings"):
                if file.endswith(".xml"):
                    if file.startswith("drawing"):
                        os.chdir("drawings")
                        drawingFilesS.append(minidom.parse(file))
                        get_ipython().run_line_magic('cd', '..')
        print(" . . . Changing directory to the student files folder . . .")
        get_ipython().run_line_magic('cd', '..')
        get_ipython().run_line_magic('cd', '..')

        stDicS = styleDictionary(stylesS)
        dxfDicS = dxfsDictionary(stylesS)
        strArrS = None
        if strS:
            strArrS = sharedString(strS)
        themeNameS = getTheme(themesS)
        sheetNamesS = sheetNames(workbookS)
        if identifyHiddenSheet(workbookS):
            hidIndices = identifyHiddenSheet(workbookS)
            temp = []
            for i,item in enumerate(worksheetsS):
                if any(i==ind for ind in hidIndices):
                    continue
                else:
                    temp.append(item)
            worksheetsS = temp
            temp = []
            for i,item in enumerate(sheetNamesS):
                if any(i==ind for ind in hidIndices):
                    continue
                else:
                    temp.append(item)
            sheetNamesS = temp        
        drawingsS = drawingFile(drawingFilesS,worksheetsS,chartFilesS)
        # dimensionS = worksheetDims(worksheetsS)
        sheetMatS = worksheetMats(worksheetsS,strArrS)
        colsWS = worksheetColsW(worksheetsS)
        hidColsS = worksheetColHidden(worksheetsS)
        freezeS = worksheetFreeze(worksheetsS)        
        conditionalFS = worksheetConditionalFormatting(worksheetsS)
        sheetProtectS = worksheetProtection(worksheetsS)
        studentFileName = studentFile.rsplit('.',1)[0]
        print("This is the student file:",studentFileName)

        res,df = finalExam(studentFileName,df)

        print(res)
    print(" . . . Change to the root folder that Jupyter is running in . . .")
    get_ipython().run_line_magic('cd', '..')
df = df.set_index('0')
df = df.T
writer = pd.ExcelWriter('ExcelResults_midterm.xlsx', engine='xlsxwriter',options={'strings_to_numbers': True})
df.to_excel(writer, sheet_name='sheet1')
writer.save()


# In[ ]:


get_ipython().run_line_magic('pwd', '')


# In[ ]:


# %cd ..


# In[ ]:


# I assumed the order of sheets is correct
# If the cell is empty, we cannot check the style, although students might have done it correctly
# Use separate name to name the columns
# I assume there is no more than one conditional formatting on each column
# If you are asking students to use a theme and you want them to resize columns, ask them to apply the best fit (auto column width) 
# It won't work if the students change the columns themeselve at first
# I assumed we don't have more than one chart on each sheet, because the order is the order of its creation (prevalidation having 2 charts)
# we cannot check the password, I think we do not need it either
# In conditional formatting we don't have the specific styles for cells, so we only can check the formula

# hw4
# adani54: numberformat 8 (q6) and numberformat 15 (q15.c)

# hw6
# 38) assign the position of the top-left part of the logo to a cell
# there should not be more than one drawing per worksheet

# hw7
# Note that using formulaF it is imperative that the teacher has a formula all over the range!

# in the code I equalized using containText and equal for the conditional formatting
# pivot tables should not start from the same cell on different worksheets
# If there are more than 9 sheets, the order would be 1, 10, 11, ..., 2, 20, 21,...


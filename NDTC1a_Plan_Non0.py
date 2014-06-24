# -*- coding: utf-8 -*-
# Welcome to the DataNitro Editor
# Use Cell(row,column).value, or Cell(name).value, to read or write to cells
# Cell(1,1) and Cell("A1") refer to the top-left cell in the spreadsheet
# 
# Note: To run this file, save it and run it from Excel (click "Run from File")
# If you have trouble saving, try using a different directory

from params import *
import numpy as np
import time

#=====================
### CUSTOM FUNCTIONS FOR EACH TEST
#=====================

# FUNCTION TO GET THE LISTS OF NAMES AND VALUES OF SOURCE CELLS
def getSrcCellsValueRange( mainSheet, totalFields, valsCol, i, j, k=0, numProducts=0 ):
    ''' Returns lists of Values and Names of the Source cells
        for the given i, j, and k indices '''
    src_RowIndex = ((k-1)*numProducts*totalFields) +\
                   ((i-1)*totalFields) + j + 1
    
    srcCells_Val = np.array( Cell( mainSheet, src_RowIndex, valsCol ).horizontal )
    srcCells = np.array( Cell( mainSheet, src_RowIndex, valsCol ).horizontal_range )
    srcCells_Name = np.array( [ cell.name for cell in srcCells ] )

    # Ignore all NoneType cells
    srcCells_Name = np.array([ srcCells_Name[i] for i,item in
                               enumerate(srcCells_Val) if item!=None])
    srcCells_Val = np.array([ float( re.sub('[^\d\.\-]','', str(i)) ) \
                              for i in srcCells_Val if i!=None])

    return srcCells_Val, srcCells_Name

# FUNCTION TO GET BAD CELLS' VALUE AND NAME USING A CRITERIA
def getBadCellsValueRange( srcCells_Val, srcCells_Name):
    ''' Returns the Values and Names of the Bad Cells based on the Criterion '''
    # (We can equivalently use either the filter method or the list comprehensions)
    # IMP: But with indexing, must convert numpy array into regular list!
    if ( len(srcCells_Val) > 0 ):
        # CRITERION #
        badCells_Val = srcCells_Val[ srcCells_Val == 0 ].tolist() 
        badCells_Range = srcCells_Name[ srcCells_Val == 0 ].tolist()
    else:
        badCells_Val = badCells_Range = []
    return badCells_Val, badCells_Range

# FUNCTION TO GET THE ADDRESSES FOR DESTINATION CELLS
def getDestCellsAdd( i, j, k, numProducts=0, totalFields=0, destScenIndex=0):
    ''' Returns Addresses for the Prod, Rate, Stat, and Value Cells in Dest. Sheet'''
    #Start with 3rd row, b/c 1st row has Test Title and 2nd row has Threshold
    dest_RowIndex = ((destScenIndex-1)*(numProducts*totalFields*2)) + \
                                 ((i-1)*(totalFields*2)) + ((j-1)*2) + 1 + 1
                
    destCells_Scen = Cell( dest_RowIndex, 1)
    destCells_Prod = Cell( dest_RowIndex, 2)
    destCells_Field = Cell( dest_RowIndex, 3)
    destCells_Stat = Cell( dest_RowIndex, 4)
    destCells_Add  = Cell( dest_RowIndex, 5)
    destCells_Val  = Cell( dest_RowIndex + 1, 5)
    return destCells_Scen, destCells_Prod, destCells_Field, destCells_Stat, \
           destCells_Add, destCells_Val

#=====================
### GENERIC FUNCTIONS FOR ALL TESTS
#=====================

# FUNCTION TO GET NAMES OF BASELINE AND SCENARIO: 
def getBaselineScenarioNames( mainSheet, scenCol, scenarioOnly ):
    """ Returns a list of Scenario and Baseline's Names"""
    scenarios = np.array( Cell( mainSheet, 2, scenCol).vertical )
    scenarios = [ str.encode('ascii') for str in scenarios ]

    seen = set()
    scenarioList = [ item for item in scenarios \
                          if item not in seen and not seen.add(item) ];
    # 'not in' operation on a set is twice as fast as on a List:
    #scenarioList = []
    #[scenarioList.append(item) for item in scenarios
    #                        if item not in scenarioList];
    numScenarios = len(scenarioList)
    
    # COMMENT OUT BELOW IF YOU WANT ACTUAL NAMES OF BASELINE AND SCENARIO
    if numScenarios == 1:
        scenarioList = ['Baseline']
    else: 
        if scenarioOnly == 1:
            scenarioList = ['Scenario']
        else:
            scenarioList = ['Baseline', 'Scenario']
    return scenarioList, numScenarios

# FUNCTION TO GET NAMES AND COUNT OF ALL PRODUCTS: 
def getProdsNamesCount( mainSheet, prodCol ):
    """ Returns a list of all Products' Names and their Count"""
    allProducts = np.array( Cell( mainSheet, 2, prodCol).vertical )
    allProducts = [ str.encode('ascii') for str in allProducts ]

    seen = set()
    prodsList = [ item for item in allProducts \
                          if item not in seen and not seen.add(item) ];
    numProducts = len(prodsList)
    return prodsList, numProducts

# FUNCTION TO GET NAMES AND COUNTS OF ALL RATE FIELDS:
def getFieldsNamesCount( mainSheet, kpisCol ):
    """ Returns a list of all Rate Fields' names and their Count"""
    allFields = Cell( mainSheet, 2, kpisCol ).vertical 
    allFields = [ str.encode('ascii') for str in allFields ]

    seen = set()
    fieldsList = [ item for item in allFields \
                   if item not in seen and not seen.add(item) ]
    totalFields = len(fieldsList)
    return fieldsList, totalFields

# FUNCTION TO CONVERT SOURCE CELLS TO FLOAT TYPE
def convertSourceToFloat( srcCells_Val ):
    ''' Returns the Source Cells converted to Float type '''
    if ( len(srcCells_Val) > 0 ):
##        time.sleep(1)
##        print srcCells_Val[0].tolist()
        if ( type( srcCells_Val[0].tolist() ) == unicode ):
            srcCells_Val = np.array( [ float( re.sub('\D', '', a) ) \
                                    for a in srcCells_Val ] )
            # replace lstrip & replace w/ regular expression: a.encode('ascii').lstrip('R ').replace(',', '')
        else:
            srcCells_Val = np.array( [ float(a) for a in srcCells_Val ] )
    return srcCells_Val

# FUNCTION TO CONVERT BAD CELLS NAMES TO DATES
def convertBadCellsToDate( badCells_Name ):
    ''' Returns a list of dates corresponding to the list of cell names '''
    badCells_Cols_1stRow = [ re.sub('\d', '', cell) + '1' \
                             for cell in badCells_Name ]
    badCells_Date = [ Cell(elem).value.strftime(' %b-%y') \
                      for elem in badCells_Cols_1stRow]
    return badCells_Date
            
# FUNCTION TO CREATE/CLEAR NDTC REPORT WORKSHEET
def createOrClearReportWorksheet(destSheet, i, j, k=0, destScenIndex=0, \
                                 scenarioOnly=0):
    """ Creates a new worksheet if not already present,
        else Clears it if the script is being run again"""
    if destSheet not in all_sheets():
        new_sheet(destSheet)
    elif scenarioOnly == 1: #if only scenario
        if j == -1:       #if only 1 parameter displayed
            if i == 1:
                clear_sheet(destSheet)
        elif i == 1 and j == 1:
            clear_sheet(destSheet)
    elif j == -1:
        if i == 1 and destScenIndex == 1:
            clear_sheet(destSheet)
    elif i == 1 and j == 1 and destScenIndex == 1:
            clear_sheet(destSheet)

# FUNCTION TO ENTER HOUSE-KEEPING DATA ONTO THE WORKSHEET
def printTitle( destSheet, testTile):
    '''Prints Title of the test and Threshold on the Sheet before values'''
    Cell(destSheet, 1, 1).value = testTitle
    Cell(destSheet, 1, 1).font.bold = True
    
    active_sheet(all_sheets()[0])

# FUNCTION TO SET THE STATUS (WITH COLORS) OF EACH RATE FIELD
def setRatesStatus( destSheet, badCells_Val, destCells_Stat, statColor):
    ''' Sets the statuses of each of the Rate Field based on
        the existence of any bad cells '''
    if len(badCells_Val) == 0:
        Cell( destSheet, destCells_Stat.name ).value = 'OK'
        Cell( destSheet, destCells_Stat.name ).font.color = statColor[0]
    else:
        Cell( destSheet, destCells_Stat.name ).value = 'WARN'
        Cell( destSheet, destCells_Stat.name ).font.color = statColor[1]
        
#=====================
### MAIN FUNCTION BODY
#=====================

def ndtc1( destSheet, testTitle, numRates, \
          scenCol, prodCol, kpisCol, valsCol, ignoreRate, scenarioOnly):

    ## MAIN SHEET: Get the name of the sheet containing data
    mainSheet = ( all_sheets()[0] ).encode('ascii')
    active_sheet(mainSheet) #MUST be included to get Source Data
    
    ## PRODUCTS: Get their count and enlist their names
    # (We have total of 16 cells representing Rates and Balances)
    prodsList, numProducts = getProdsNamesCount( mainSheet, prodCol )

    ## FIELDS: Get their count and enlist their names as strings 
    fieldsList, totalFields = getFieldsNamesCount( mainSheet, kpisCol )

    ## SCENARIOS: Get names of baseline and scenario
    if (scenCol==0):
        scenarioList = ['Analyze']
        numScenarios = 1
    else:
        scenarioList, numScenarios = getBaselineScenarioNames( mainSheet, scenCol, \
                                                           scenarioOnly)
    
    ## SET COLORS: for products, rates, and statuses
    scenColor = prodColor = fieldsColor = ['blue', 'black']
    statColor = ['green', 'red']

    # INDEX FOR PRINTING DEST. SCENARIO CELLS
    destScenIndex = 1
    
    ## FOR EACH SCENARIO:
    if (scenarioOnly == 0):
        kstart = 1
    else:
        kstart = 2

    for k in range( kstart, numScenarios+1):
        
        ## FOR EACH PRODUCT:
        for i in range(1, numProducts+1):
            
            ## FOR EACH OF THE RATE'S & BALANCE'S ATTRIBUTES:
            for j in range(1, totalFields+1):
                
                ## SOURCE CELLS: Get Values and Range of the Corresponding Cells
                srcCells_Val, srcCells_Name = getSrcCellsValueRange( mainSheet, \
                                                                     totalFields, \
                                                                     valsCol, \
                                                                     i, j, k, numProducts)

                ## BALANCE TYPE' CONVERSION: First convert all Balances into Float type
                srcCells_Val = convertSourceToFloat( srcCells_Val )
                
                ## BAD CELLS INFO: Values and Names/Dates of the Outlier Cells
                badCells_Val, badCells_Name = getBadCellsValueRange( srcCells_Val, \
                                                                      srcCells_Name)
                badCells_Date = convertBadCellsToDate( badCells_Name)
                
                ## DESTINATION SHEET GENERATION:
                createOrClearReportWorksheet( destSheet, i, j, k, destScenIndex, scenarioOnly)
                printTitle( destSheet, testTitle )
                
                ## DESTINATION CELLS
                destCells_Scen, destCells_Prod, destCells_Field, destCells_Stat, \
                            destCells_Add, destCells_Val = getDestCellsAdd( i, j, 0, \
                                                            numProducts, totalFields, \
                                                            destScenIndex)

                ## POPULATE DESTINATION CELLS:
                if ( i == 1 and j == 1 ): #Scenario Name printed only once
                    Cell( destSheet, destCells_Scen.name ).value = scenarioList[k-1]
                    Cell( destSheet, destCells_Scen.name ).font.color = scenColor[i%2]
                                 
                if ( j == 1 ):  #Product Name printed only once
                    Cell( destSheet, destCells_Prod.name ).value = prodsList[i-1]
                    Cell( destSheet, destCells_Prod.name ).font.color = prodColor[i%2]
                
                Cell( destSheet, destCells_Field.name ).value = fieldsList[j-1]
                Cell( destSheet, destCells_Field.name ).font.color = fieldsColor[j%2]

                if ( fieldsList[j-1] not in ignoreRate ):
                    setRatesStatus( destSheet, badCells_Val, destCells_Stat, statColor);

                    Cell( destSheet, destCells_Add.name ).horizontal = badCells_Date
                    Cell( destSheet, destCells_Val.name ).horizontal = badCells_Val

        # Increment the index for Scenario Destination Cell 
        destScenIndex += 1

#=====================
### SET PARAMETERS AND RUN
#=====================
if __name__ == '__main__':
    
    destSheet = 'NDTC1a'
    testTitle = 'All Rates and Balances are populated for All Products i.e. Non-Zero'\
                'under Baseline and Scenario'
    ignoreRate = ['Rate Discretion', 'Liquidity Premium']
    scenarioOnly = 0  #1 for Scenario only, 0 for Both Scenario and Baseline

    ndtc1( destSheet, testTitle, numRates, scenCol, prodCol, kpisCol, valsCol, \
          ignoreRate, scenarioOnly )
    display_sheet( destSheet)

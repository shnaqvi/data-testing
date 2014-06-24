# -*- coding: utf-8 -*-
# Welcome to the DataNitro Editor
# Use Cell(row,column).value, or Cell(name).value, to read or write to cells
# Cell(1,1) and Cell("A1") refer to the top-left cell in the spreadsheet
# 
# Note: To run this file, save it and run it from Excel (click "Run from File")
# If you have trouble saving, try using a different directory

import numpy as np

#=====================
### CUSTOM FUNCTIONS FOR EACH TEST
#=====================

# FUNCTION TO GET THE LISTS OF NAMES AND VALUES OF SOURCE CELLS
def getSrcCellsValueRange( mainSheet, totalFields, valsCol, i, j=0, k=0, \
                           numProducts=0, numRates=0, begEndBalInd=[], \
                           spreadFieldsInd=[] ):
    ''' Returns lists of Values and Names of the Source cells
        for the given i and j indices '''
    src_RowIndices = np.zeros((2,1), dtype=int)
    srcCells_Val = np.zeros( (2, len( Cell( mainSheet, 2, valsCol).horizontal)))

    for row in range(0,2):
        src_RowIndices[row]= row*(numProducts*totalFields) + ((i-1)*totalFields) + \
                             numRates + (begEndBalInd[1]+1) + 1
        srcCells_Val[row] = Cell( mainSheet, np.asscalar( src_RowIndices[row][0]),\
                                  valsCol).horizontal
    srcCells = Cell( mainSheet, np.asscalar( src_RowIndices[1][0]), \
                     valsCol ).horizontal_range
    srcCells_Name = np.array( [ cell.name for cell in srcCells ] )
    return srcCells_Val, srcCells_Name

# FUNCTION TO GET BAD CELLS' VALUE AND NAME USING A CRITERIA
def getBadCellsValueRange( srcCells_Val, srcCells_Name, thresh, bpsVar):
    ''' Returns the Values and Names of the Bad Cells based on the Criterion '''
    # (We can equivalently use either the filter method or the list comprehensions)
    # IMP: But with indexing, must convert numpy array into regular list!
    # NOTE: Division must be with an absolute value to obtain true increase in neg. values as well

    if ( len(srcCells_Val) > 0 ):
        # CRITERION # (division by 0 in float results in float(inf))
        var = (srcCells_Val[1] - srcCells_Val[0])/srcCells_Val[0] * 100 
        var = np.array([ round(num, 1) for num in var])

        if ( bpsVar>0 ):
            badCells_Var = var[ var>thresh ].tolist()
            badCells_Val = srcCells_Val[1][ var>thresh ].tolist()
            badCells_Name = srcCells_Name[ var>thresh ].tolist()
        else:
            badCells_Var = var[ var<-thresh ].tolist()
            badCells_Val = srcCells_Val[1][ var<-thresh ].tolist()
            badCells_Name = srcCells_Name[ var<-thresh ].tolist()

    else:
        badCells_Var = badCells_Val = badCells_Name = []

    return badCells_Var, badCells_Val, badCells_Name

# FUNCTION TO GET THE ADDRESSES FOR DESTINATION CELLS
def getDestCellsAdd( i, j, k, numProducts=0, totalFields=0, numBalance=0, \
                     destScenIndex=0, begEndBalInd=[], spreadFieldsInd=[] ):
    ''' Returns Addresses for the Prod, Rate, Stat, and Value Cells in Dest. Sheet'''
    #Start with 3rd row, b/c 1st row has Test Title and 2nd row has Threshold
    dest_RowIndex = ((destScenIndex-1)*(numProducts*3)) +\
                    ((i-1)*3) + 1 + 1 + 1
    
    destCells_Scen = Cell( dest_RowIndex, 1)
    destCells_Prod = Cell( dest_RowIndex, 2)
    destCells_Rate = Cell( dest_RowIndex, 3)
    destCells_Stat = Cell( dest_RowIndex, 4)
    destCells_Add  = Cell( dest_RowIndex, 5)
    destCells_Val  = Cell( dest_RowIndex + 1, 5)
    destCells_Var  = Cell( dest_RowIndex + 2, 5)
    return destCells_Scen, destCells_Prod, destCells_Rate, destCells_Stat, \
           destCells_Add, destCells_Val, destCells_Var

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
    # 'not in' operation on a set is twice as fast as on a List:
    #prodsList = []
    #[prodsList.append(item) for item in allProducts
    #                        if item not in prodsList];
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

# FUNCTION TO GET NAMES AND COUNTS OF ALL BALANCE FIELDS:
def getBalanceNamesCount( mainSheet, kpisCol, numRates, numBalance ):
    """ Returns a list of all Rate Fields' names and their Count"""
    balanceList = CellRange( mainSheet, (2+numRates,kpisCol), \
                             (2+numRates+numBalance-1,kpisCol) ).value
    balanceList = [ str.encode('ascii') for str in balanceList ]
    return balanceList

# FUNCTION TO GET THE LIST OF INDICES FOR A PRODUCT'S BEGINNING & ENDING BALANCES
def getBegEndBalIndices( fieldsList ):
    ''' Returns the indices for the Beg & End. Balance field attributes '''
    fields = ['beginning balance', 'ending balance']
    indices = []
    for j in range(0, len(fields) ):
        indices.extend( [i for i,item in enumerate( fieldsList ) \
                       if fields[j] in item.lower()] )
    return indices    

# FUNCTION TO CONVERT SOURCE CELLS TO FLOAT TYPE
def convertSourceToFloat( srcCells_Val ):
    ''' Returns the Source Cells converted to Float type '''
    if ( len(srcCells_Val) > 0 ):
        if ( type( srcCells_Val[0].tolist() ) == unicode ):
            srcCells_Val = np.array( [ float( re.sub('\D', '', a) ) \
                                    for a in srcCells_Val ] )
            # replace lstrip & replace w/ regular expression: a.encode('ascii').lstrip('R ').replace(',', '')
        else:
            for i in range(0, len(srcCells_Val)):
                srcCells_Val[i] = np.array( [ float(a) for a in srcCells_Val[i] ] )
    return srcCells_Val

# FUNCTION TO CONVERT BAD CELLS NAMES TO DATES
def convertBadCellsToDate( badCells_Name ):
    ''' Returns a list of dates corresponding to the list of cell names '''
    badCells_Cols = [ re.sub('\d', '', cell) for cell in badCells_Name ]
    badCells_1stRowCols = [ elem + '1' for elem in badCells_Cols ]
    badCells_Date = [ Cell(elem).value.strftime(' %b-%y') \
                      for elem in badCells_1stRowCols]
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

# FUNCTION TO GET THESHOLD FROM THE WORKSHEET
def getThresholdFromSheet( destSheet, thresh, bpsVar ):
    '''Gets the User-entered Threshold value before Clearing the Sheet'''
    if destSheet in all_sheets():
        if Cell( destSheet, 2, 2) != None:
            thresh = Cell( destSheet, 2, 2).value
        if Cell( destSheet, 2, 4) != None:
            bpsVar = Cell( destSheet, 2, 4).value
    return thresh, bpsVar

# FUNCTION TO ENTER HOUSE-KEEPING DATA ONTO THE WORKSHEET
def printTitleAndThreshold( destSheet, testTile, thresh=None, bpsVar=None):
    '''Prints Title of the test and Threshold on the Sheet before values'''
    Cell(destSheet, 1, 1).value = testTitle
    Cell(destSheet, 1, 1).font.bold = True

    if thresh != None:
        Cell(destSheet, 2, 1).value = 'Perc_Inc:'
        Cell(destSheet, 2, 1).font.bold = True

        Cell(destSheet, 2, 2).value = thresh
        Cell(destSheet, 2, 2).color = 'yellow'
        
    if bpsVar != None:
        Cell(destSheet, 2, 3).value = 'MktSpd_Var:'
        Cell(destSheet, 2, 3).font.bold = True

        Cell(destSheet, 2, 4).value = bpsVar
        Cell(destSheet, 2, 4).color = 'yellow'

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

def main( destSheet, testTitle, numRates, numBalance,\
          kpisCol, prodsCol, valsCol, ignoreRate, scenarioOnly, thresh, bpsVar):

    ## MAIN SHEET: Get the name of the sheet containing data & Set it as active
    mainSheet = ( all_sheets()[0] ).encode('ascii')
    active_sheet(mainSheet)

    # THRESHOLD EXTRACTION FROM SHEET
    thresh, bpsVar = getThresholdFromSheet( destSheet, thresh, bpsVar)
    
    ## PRODUCTS: Get their count and enlist their names
    # (We have total of 16 cells representing Rates and Balances)
    prodsList, numProducts = getProdsNamesCount( mainSheet, prodCol )

    ## FIELDS: Get their count and enlist their names as strings 
    fieldsList, totalFields = getFieldsNamesCount( mainSheet, kpisCol )
    
    ## BALANCE: Get their count and enlist their names as strings 
    balanceList = getBalanceNamesCount( mainSheet, kpisCol, numRates, numBalance )

    ## BEG. END. BAL. INDICES: Get Indices for Beg. & End. Bal. attributes
    begEndBalInd = getBegEndBalIndices( balanceList )
    
    ## SCENARIOS: Get names of baseline and scenario
    scenarioList, numScenarios = getBaselineScenarioNames( mainSheet, scenCol, \
                                                           scenarioOnly)
    
    ## SET COLORS: for products, rates, and statuses
    scenColor = prodColor = balanceColor = ['blue', 'black']
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
            
            ## SOURCE CELLS: Get Values and Range of the Corresponding Cells
            srcCells_Val, srcCells_Name = getSrcCellsValueRange( mainSheet, \
                                                totalFields, valsCol, i, 0, k, \
                                                numProducts, numRates, \
                                                begEndBalInd, [])
            
            ## BALANCE TYPE' CONVERSION: First convert all Balances into Float type
            srcCells_Val = convertSourceToFloat( srcCells_Val )
            
            ## BAD CELLS INFO: Values and Names of the Outlier Cells
            badCells_Var, badCells_Val, badCells_Name = getBadCellsValueRange( srcCells_Val, \
                                                            srcCells_Name, thresh, bpsVar)
            badCells_Date = convertBadCellsToDate( badCells_Name)
            
            ## DESTINATION SHEET GENERATION & THRESHOLD EXTRACTION & PRINTING: (j=-1 if j not used)
            createOrClearReportWorksheet( destSheet, i, -1, 0, \
                                          destScenIndex, scenarioOnly)
            printTitleAndThreshold( destSheet, testTitle, thresh, bpsVar)
            
            ## DESTINATION CELLS
            destCells_Scen, destCells_Prod, destCells_Rate, destCells_Stat, \
                    destCells_Add, destCells_Val, destCells_Var = \
                    getDestCellsAdd( i, 0, 0, numProducts, 0, numBalance, \
                                     destScenIndex, begEndBalInd, [])

            ## POPULATE DESTINATION CELLS:
            if ( i == 1 ): #Scenario Name printed only once
                Cell( destSheet, destCells_Scen.name ).value = scenarioList[destScenIndex-1]
                Cell( destSheet, destCells_Scen.name ).font.color = scenColor[i%2]

            Cell( destSheet, destCells_Prod.name ).value = prodsList[i-1]
            Cell( destSheet, destCells_Prod.name ).font.color = prodColor[i%2]
            
            Cell( destSheet, destCells_Rate.name ).value = balanceList[begEndBalInd[1]]
            Cell( destSheet, destCells_Rate.name ).font.color = balanceColor[1]

            setRatesStatus( destSheet, badCells_Val, destCells_Stat, statColor);

            Cell( destSheet, destCells_Add.name ).horizontal = badCells_Date
            Cell( destSheet, destCells_Val.name ).horizontal = badCells_Val

            Cell( destSheet, destCells_Var.name ).horizontal = badCells_Var
            CellRange( destSheet, (destCells_Var.row, destCells_Var.col), \
                        (destCells_Var.row, destCells_Var.col+len(badCells_Var)) ).font.color = 'red'
            
        # Increment the index for Scenario Destination Cell 
        destScenIndex += 1
    
#=====================
### SET PARAMETERS AND RUN
#=====================
numRates = 3
numBalance = 10
scenCol = 2
prodCol = 4
kpisCol = 5
valsCol = 6
destSheet = 'NDTC15a'
testTitle = "Ending Balances are Not > x% higher in Scenario than in Baseline "\
            "for all products with Market Spend Increased by ybps"
ignoreRate = 'Rate Discretion'
scenarioOnly = 1  #MUST BE 1 FOR COMPARING SCENARIO W/ BASELINE: 1 for Scenario 
thresh = 30
bpsVar = 10

main( destSheet, testTitle, numRates, numBalance,\
      kpisCol, prodCol, valsCol, ignoreRate, scenarioOnly, thresh, bpsVar)
display_sheet( destSheet)

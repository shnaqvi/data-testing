
##NOTE: If we need to do computation on any individual field, either do a "MoM
# analysis" on it or comapre it with itself in the scenario, then we can define
# it outside of the main body and pass it as a parameter.
#BUT IF that field needs to be compared with other fields in the same month,
# then we need to include a Specific function to get indices of all the fields
# that need to go into the computation

from params import *
import numpy as np

#=====================
### CUSTOM FUNCTIONS FOR EACH TEST
#=====================

# FUNCTION TO GET THE LISTS OF NAMES AND VALUES OF SOURCE CELLS
def getSrcCellsValueRange( mainSheet, totalFields, valsCol, i, j=0, k=0, \
                           numProducts=0, numRates=0, specFieldsInd=[], \
                           spreadFieldsInd=[] ):
    ''' Returns lists of Values and Names of the Source cells
        for the given i and j indices '''
    src_RowIndex = ((k-1)*numProducts*totalFields) + ((i-1)*totalFields) + \
                   (specFieldsInd[j-1]+1) + 1
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
def getBadCellsValueRange( srcCells_Val, srcCells_Name, thresh ):
    ''' Returns the Values and Names of the Bad Cells based on the Criterion '''
    # (We can equivalently use either the filter method or the list comprehensions)
    # IMP: But with indexing, must convert numpy array into regular list!
    # NOTE: Division must be with an absolute value to obtain true increase in neg. values as well

    if ( len(srcCells_Val) > 0 ):
        # NEW LISTS FOR MONTH OVER MONTH COMPARISON
        srcCells_Val_NextMo = srcCells_Val[1:]
        srcCells_Val_LastMo = srcCells_Val[:len(srcCells_Val)-1]
        srcCells_Val_New = srcCells_Val_NextMo[:]
        srcCells_Name_New = srcCells_Name[1:]
        
        # CRITERION # (division by 0 in float results in float(inf))
        var = np.true_divide((srcCells_Val_NextMo - srcCells_Val_LastMo),abs(srcCells_Val_LastMo)) * 100
        var = np.array([ round(num, 2) for num in var])

        badCells_Var = var[ abs(var)> thresh ].tolist()
        badCells_Val = srcCells_Val_New[ abs(var) > thresh ].tolist()
        badCells_Name = srcCells_Name_New[ abs(var)> thresh].tolist()
        
    else:
        badCells_Var = badCells_Val = badCells_Name = []

    return badCells_Var, badCells_Val, badCells_Name

# FUNCTION TO GET THE ADDRESSES FOR DESTINATION CELLS
def getDestCellsAdd( i, j, k, numProducts=0, totalFields=0, numBalance=0, \
                     destScenIndex=0):
    ''' Returns Addresses for the Prod, Rate, Stat, and Value Cells in Dest. Sheet'''
    #Start with 3rd row, b/c 1st row has Test Title and 2nd row has Threshold
    dest_RowIndex = ((destScenIndex-1)*(numProducts*numBalance*3)) +\
                    ((i-1)*numBalance*3) + ((j-1)*3) + 1 + 1 + 1
    
    destCells_Scen = Cell( dest_RowIndex, 1)
    destCells_Prod = Cell( dest_RowIndex, 2)
    destCells_Field = Cell( dest_RowIndex, 3)
    destCells_Stat = Cell( dest_RowIndex, 4)
    destCells_Add  = Cell( dest_RowIndex, 5)
    destCells_Val  = Cell( dest_RowIndex + 1, 5)
    destCells_Var  = Cell( dest_RowIndex + 2, 5)
    return destCells_Scen, destCells_Prod, destCells_Field, destCells_Stat, \
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
def getBalanceNamesCount( fieldsList, numRates, numBalance ):
    """ Returns a list of all Rate Fields' names and their Count"""
    balanceList = fieldsList[numRates:numRates+numBalance+1]
    return balanceList

## FUNCTION TO GET THE LIST OF INDICES FOR SPECIFIC FIELDS
def getSpecFieldsIndices( fieldsList, specFields ):
    """ Returns the indices for the Specific Fields"""
    fieldsList = np.char.strip(fieldsList) #strip leading spaces

    indices = [ i for field in specFields for i,item in enumerate(fieldsList)\
               if re.search(field, item, re.IGNORECASE) ]
    numSpecFields = len(specFields)
    
    return indices, numSpecFields      

# FUNCTION TO CONVERT SOURCE CELLS TO FLOAT TYPE
def convertSourceToFloat( srcCells_Val ):
    ''' Returns the Source Cells converted to Float type '''
    if ( len(srcCells_Val) > 0 ):
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
def getThresholdFromSheet( destSheet, thresh ):
    '''Gets the User-entered Threshold value before Clearing the Sheet'''
    if destSheet in all_sheets():
        if Cell( destSheet, 2, 2) != None:
            thresh = Cell( destSheet, 2, 2).value
    return thresh

# FUNCTION TO ENTER HOUSE-KEEPING DATA ONTO THE WORKSHEET
def printTitleAndThreshold( destSheet, testTile, thresh=None ):
    '''Prints Title of the test and Threshold on the Sheet before values'''
    Cell(destSheet, 1, 1).value = testTitle
    Cell(destSheet, 1, 1).font.bold = True

    if thresh != None:
        Cell(destSheet, 2, 1).value = 'Threshold:'
        Cell(destSheet, 2, 1).font.bold = True

        Cell(destSheet, 2, 2).value = thresh
        Cell(destSheet, 2, 2).color = 'yellow'
        
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

def ndtc12( destSheet, testTitle, numRates, numBalance,\
          kpisCol, prodsCol, valsCol, ignoreRate, scenarioOnly, thresh, specFields):

    ## MAIN SHEET: Get the name of the sheet containing data & Set it as active
    mainSheet = ( all_sheets()[0] ).encode('ascii')
    active_sheet(mainSheet)

    # THRESHOLD EXTRACTION FROM SHEET
    thresh = getThresholdFromSheet( destSheet, thresh)
    
    ## PRODUCTS: Get their count and enlist their names
    # (We have total of 16 cells representing Rates and Balances)
    prodsList, numProducts = getProdsNamesCount( mainSheet, prodCol )

    ## FIELDS: Get their count and enlist their names as strings 
    fieldsList, totalFields = getFieldsNamesCount( mainSheet, kpisCol )
    
    ## BALANCE: Get their count and enlist their names as strings 
    balanceList = getBalanceNamesCount( fieldsList, numRates, numBalance )

    ## SPECIFIC FIELDS: Get Indices of specific fields
    specFieldsInd, numSpecFields = getSpecFieldsIndices( fieldsList, specFields )
    numBalance = numSpecFields

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

            ## FOR EACH BALANCE FIELD:
            for j in range(1, numBalance+1):
            
                ## SOURCE CELLS: Get Values and Range of the Corresponding Cells
                srcCells_Val, srcCells_Name = getSrcCellsValueRange( mainSheet, \
                                                    totalFields, valsCol, i, j, k, \
                                                    numProducts, numRates, \
                                                    specFieldsInd, [])
                
                ## BALANCE TYPE' CONVERSION: First convert all Balances into Float type
                srcCells_Val = convertSourceToFloat( srcCells_Val )
                
                ## BAD CELLS INFO: Values and Names of the Outlier Cells
                badCells_Var, badCells_Val, badCells_Name = getBadCellsValueRange( srcCells_Val, \
                                                                srcCells_Name, thresh)
                badCells_Date = convertBadCellsToDate( badCells_Name)
                
                ## DESTINATION SHEET GENERATION & THRESHOLD EXTRACTION & PRINTING: (j=-1 if j not used)
                createOrClearReportWorksheet( destSheet, i, j, 0, \
                                              destScenIndex, scenarioOnly)
                printTitleAndThreshold( destSheet, testTitle, thresh )
                
                ## DESTINATION CELLS
                destCells_Scen, destCells_Prod, destCells_Field, destCells_Stat, \
                                destCells_Add, destCells_Val, destCells_Var = \
                                getDestCellsAdd( i, j, 0, numProducts, 0, numBalance, \
                                                 destScenIndex)

                ## POPULATE DESTINATION CELLS:
                if ( i == 1 and j == 1 ): #Scenario Name printed only once
                    Cell( destSheet, destCells_Scen.name ).value = scenarioList[k-1]
                    Cell( destSheet, destCells_Scen.name ).font.color = scenColor[i%2]

                if ( j == 1): 
                    Cell( destSheet, destCells_Prod.name ).value = prodsList[i-1]
                    Cell( destSheet, destCells_Prod.name ).font.color = prodColor[i%2]

                Cell( destSheet, destCells_Field.name ).value = fieldsList[specFieldsInd[j-1]]
                Cell( destSheet, destCells_Field.name ).font.color = fieldsColor[j%2]

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
if __name__ == '__main__':

    destSheet = 'NDTC12a'
    testTitle = "Beginning & Ending Balance Must Not vary more than +/-15% Monthly "\
                "for all products in Baseline & Scenario" 
    ignoreRate = 'Rate Discretion'
    scenarioOnly = 0  #1 for Scenario only, 0 for Both Scenario and Baseline
    thresh = 15
    specFields = ['beg.*balance$', 'end.*balance$']

    ndtc12( destSheet, testTitle, numRates, numBalance,\
          kpisCol, prodCol, valsCol, ignoreRate, scenarioOnly, thresh, specFields)
    display_sheet( destSheet)

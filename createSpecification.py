#!/usr/bin/env python
"""
    Config File Creator - an important script to create Config files from a D-spec file.

    This script reads in and validates a dashboard spec excel file. It identifies
    the presence and dimensions of active question boxes within the excel file, and
    imports the data from these question boxes. After validating the imported data,
    it outputs the imported data to a .dcc.yaml file.

    You might find the xlrd documentation extremely useful for understanding parts of this script:
    http://www.lexicon.net/sjmachin/xlrd.html

    Changes in progress:
        - in import_qboxes, change return type to dictionary for information collected for each question
        - implement yaml file output function
        - Add exception handling
        - Add optparse
        - Update script to conform to the following use cases:
              $ createSpecification.py
                  - print help
              $ createSpecification.py -d <dashboard_xls_filename>
                  - generate a yaml file template, populate all sections except for the header
              $ createSpecification.py -d <dashboard_xls_filename> -c
                  - prompt user to enter details needed for yaml header, then generate yaml file.
              $ createSpecification.py -d <dashboard_xls_filename> -x <desparsed_csv_filename>
                  - in addition to the functionality of -c, alter the output yaml to reflect
                    ->the presence of empty 9999 cells within the desparses csv file.
"""
import sys, os
import xlrd, re
import unicodedata

# TODO: Remove after testing
PRINT_DEBUG = False
PRINT_IMPORT = False
PRINT_WARNINGS = False

def import_qboxes(worksheet, qboxDimensions):

    qboxHeaders, qboxFooters = zip(*qboxDimensions)

    questionNumber = 0
    totalQuestions = str(len(qboxHeaders))
    questionData = []

    # for each question:
    for i, headerRow in enumerate(qboxHeaders):
        footerRow = qboxFooters[i]
        
        question = []
        questionNumber += 1

        if PRINT_IMPORT:
            print "question #: " + str(questionNumber) + " of " + totalQuestions + "\r"

        # import the variable name
        variableName = get_variable_name(worksheet, headerRow)

        # import the Dashboard Label
        dashboardLabel = get_dashboard_label(worksheet, headerRow)

        # import that question's response values
        responseValues = get_response_values(worksheet, headerRow, footerRow)

        question.append(questionNumber)
        question.append(variableName)
        question.append(dashboardLabel)
        question.append(responseValues)

        if PRINT_IMPORT:
            print "Response Values:\r"
            for val in responseValues:
                print val

        # import that question's net numbers
        netNumbers = get_net_numbers(worksheet, headerRow, footerRow)

        if PRINT_IMPORT:
            print "Net Numbers:\r"
            for val in netNumbers:
                print val

        question.append(netNumbers)

        # check amount of response values == amount of net numbers
        if len(netNumbers) != len(responseValues):
            print "the number of reponse value entries does not match the number of net number entries for question " + str(questionNumber)
            print "please check this question in the dashboard specification form and then try this program again."
            sys.exit(0)                              # TODO: replace with proper exception handling

        # import the question's dashboard net labels.
        netLabels = get_net_labels(worksheet, headerRow, footerRow)

        question.append(netLabels)

        if PRINT_IMPORT:
            print "Net Labels:\r"
            for string in netLabels:
                print string

        # check amount of net labels == amount of netting categories
        if len(netLabels) != max(netNumbers):
            print "There are more unique net labels than there are unique net numbers!"
            print "please check question " + str(questionNumber) + " in the dashboard specification form then try this program again."
            sys.exit(0)                              # TODO: replace with proper exception handling

        # get custom netting name for this question
        nettingName = get_netting_name(responseValues, netNumbers)
        if nettingName == "none":
            nettingName = str("Question" + str(questionNumber) + "Netting")

        question.append(nettingName)

        if PRINT_IMPORT:
            print "Netting Name: " + nettingName + "\r"

        questionData.append(question)

    return questionData

def get_dashboard_label(worksheet, headerRow):

    dashboardLabel = ""

    if (worksheet.cell_type(headerRow, 4) == 0):
        print "found a blank dashboard label."
        print "please ensure the dashboard label at row " + str(headerRow+1) + " exists and then try this program again."
        sys.exit(0)                              # TODO: replace with proper exception handling
    elif (worksheet.cell_type(headerRow, 4) == 1):
            variableName = unicodedata.normalize('NFKD', worksheet.cell_value(headerRow, 4)).encode('ascii', 'ignore')
    elif (worksheet.cell_type(headerRow, 4) == 2):
        variableName = str(worksheet.cell_value(headerRow, 4))
    else:
        print "could not read a dashboard label cell."
        print "please ensure the dashboard label at row " + str(headerRow+1) + " exists and then try this program again."
        sys.exit(0)                              # TODO: replace with proper exception handling

    return variableName

def get_variable_name(worksheet, headerRow):

    variableName = ""

    if (worksheet.cell_type(headerRow, 2) == 0):
        print "found a blank variable name."
        print "please ensure the variable name at row " + str(headerRow+1) + " exists and then try this program again."
        sys.exit(0)                              # TODO: replace with proper exception handling
    elif (worksheet.cell_type(headerRow, 2) == 1):
            variableName = unicodedata.normalize('NFKD', worksheet.cell_value(headerRow, 2)).encode('ascii', 'ignore')
    elif (worksheet.cell_type(headerRow, 2) == 2):
        variableName = str(worksheet.cell_value(headerRow, 2))
    else:
        print "could not read a variable name cell."
        print "please ensure the question variable name at row " + str(headerRow+1) + " exists and then try this program again."
        sys.exit(0)                              # TODO: replace with proper exception handling

    return variableName
    

def get_netting_name(responseValues, netNumbers):

    nettingName = "none"

    # No Netting
    no_net = True
    for i in range(0, len(netNumbers)):
        if not netNumbers[i] == i+1:
            no_net = False

    # TopVersusRestUpTo10
    tvrut10 = True
    for i in range(1, len(netNumbers)):
        if not netNumbers[i] == 2:
            tvrut10 = False

    # TopTwoVersusRestUpTo10
    t2vrut10 = True
    if (len(netNumbers) >= 3) and (netNumbers[0] == 1 and netNumbers[1] == 1):
        for i in range(2, len(netNumbers)):
            if not netNumbers[i] == 2:
                t2vrut10 = False
    else:
        t2vrut10 = False

    # SixToThree
    stt = False
    if len(netNumbers) == 6:
        if netNumbers[0]==1 and netNumbers[1]==1 and netNumbers[2]==2 and netNumbers[3]==2 and netNumbers[4]==3 and netNumbers[5]==3:
            stt = True


    if no_net:
        nettingName = "0"
    elif tvrut10:
        nettingName = "TopVersusRestUpTo10"
    elif t2vrut10:
        nettingName = "TopTwoVersusRestUpTo10"
    elif stt:
        nettingName = "SixToThree"

    # TODO: Add Custom Netting Label Support for Pre-Existing Net Labels
    # (Program is functional without this, but human readers of the .dcc.yaml file might get upset if this section is neglected for long)

    return nettingName

# Input: a qbox header and footer
# Output: a list of net labels for the specified question
def get_net_labels(worksheet, headerRow, footerRow):

    netLabels = []

    # net labels start two rows below the header row
    netLabelRow = headerRow+2
    prev_label = ""
    
    # from netLabelRow to footer:
    for row in range(netLabelRow, footerRow):

        if (worksheet.cell_type(row, 5) == 0):
            print "found a blank net label entry."
            print "please ensure the net label at row " + str(row+1) + " exists and then try this program again."
            sys.exit(0)                              # TODO: replace with proper exception handling
        elif (worksheet.cell_type(row, 5) == 1):
            # get candidate net label
            label = unicodedata.normalize('NFKD', worksheet.cell_value(row, 5)).encode('ascii', 'ignore')
            
            # first net label is automatically accepted
            if len(netLabels)==0:
                prev_label = label
                netLabels.append(label)
            else:
                # all other net labels are only accepted if they differ from the netLabel before it
                if prev_label == "":
                    print "logic error in get_net_labels function."
                    sys.exit(0)                      # TODO: replace with proper exception handling
                elif label != prev_label:
                    netLabels.append(label)
                    prev_label = label
                else:
                    prev_label = label
        else:
            print "could not read a net label cell."
            print "please ensure the net label at row " + str(row+1) + " makes sense and then try this program again."
            sys.exit(0)                              # TODO: replace with proper exception handling

    return netLabels

# Input: a qbox header and footer
# Output: a list of net numbers for the specified question
def get_net_numbers(worksheet, headerRow, footerRow):

    netNumbers = []

    # net numbers start two rows below the header row
    netNumberRow = headerRow+2
    
    # from netNumberRow to footer:
    for row in range(netNumberRow, footerRow):

        if (worksheet.cell_type(row, 4) == 0):
            print "found a blank net number entry."
            print "please ensure the net number at row " + str(row+1) + " exists and then try this program again."
            sys.exit(0)                              # TODO: replace with proper exception handling
        elif (worksheet.cell_type(row, 4) == 1):
            print "found a net number entry which was not a number: " + \
            unicodedata.normalize('NFKD', worksheet.cell_value(row, 4)).encode('ascii', 'ignore')
            print "please ensure the net number at row " + str(row+1) + " is a number and then try this program again."
            sys.exit(0)                              # TODO: replace with proper exception handling
        elif (worksheet.cell_type(row, 4) == 2):
            val = worksheet.cell_value(row, 4)
            netNumbers.append(val)
        else:
            print "could not read a net number cell."
            print "please ensure the net number at row " + str(row+1) + " exists and then try this program again."
            sys.exit(0)                              # TODO: replace with proper exception handling

    return netNumbers

# Input: a qbox header and footer
# Output: a list of response values for the specified question
def get_response_values(worksheet, headerRow, footerRow):

    responseValues = []
    expected_value = 1.0

    # response values start two rows below the header row
    responseValueRow = headerRow+2
    
    # from responseValueRow to footer:
    for row in range(responseValueRow, footerRow):

        if (worksheet.cell_type(row,1) == 0):
            if PRINT_WARNINGS:
                print "found a blank cell where a response value entry was expected."
                print "(expected response value of " + str(expected_value) + " at row #" + str(row+1) + ")\r"
                print "adding expected reponse value instead."
            responseValues.append(expected_value)
        elif (worksheet.cell_type(row, 1) == 1):
            if PRINT_WARNINGS:
                print "found a response value entry which was not a number: " + \
                unicodedata.normalize('NFKD', worksheet.cell_value(row, 1)).encode('ascii', 'ignore')
                print "(expected response value of " + str(expected_value) + " at row #" + str(row+1) + ")\r"
                print "adding expected response value instead."
            responseValues.append(expected_value)
        elif (worksheet.cell_type(row, 1) == 2):
            val = worksheet.cell_value(row, 1)
            if val != expected_value:
                if PRINT_WARNINGS:
                    print "response value expected was " + str(expected_value) + " and the response value received was " + str(val)
                if val == (expected_value - 1):
                    if PRINT_WARNINGS:
                        print "response values for this question appear to start at 0 instead of 1."
                        print "this program will perform as if these values start at 1."
                    responseValues.append(expected_value)
                else:
                    if PRINT_WARNINGS:
                        "using expected response value instead."
                    responseValues.append(expected_value)
            else:
                responseValues.append(expected_value)
        else:
            print "could not read a response value cell."
            print "(expected response value of " + str(expected_value) + " at row #" + str(row+1) + ")\r"
            print "adding expected reponse value instead."
            responseValues.append(expected_value)

        expected_value += 1.0

    return responseValues

# Prints the questions within each qbox header given as input
def print_qbox_questions(worksheet, qboxHeaderRows):

    for qbox_header_row in qboxHeaderRows:
        print "    " + worksheet.cell_value(qbox_header_row, 2) + ":\r"
        print "        name: \"" + worksheet.cell_value(qbox_header_row, 4) + "\"\r"

    return 0

# A useful debugging function which prints qbox data to the screen
def print_qboxes(worksheet, qboxDimensions):

    qbox_headers, qbox_footers = zip(*qboxDimensions)

    for i, qbox_header_row in enumerate(qbox_headers):
        qbox_footer_row = qbox_footers[i]    
       
        print "\r################################"
        print "QBOX " + str(i+1) + ":"
        print "Dimensions: (" + str(qbox_header_row) + ", " + str(2) + ") -- (" + str(qbox_footer_row) + ", " + str(5) + ")"

        for curr_row_num in range(qbox_header_row-1, qbox_footer_row+1):
            if curr_row_num == qbox_header_row-1:
                print "\r[Title Row " + str(curr_row_num+1) + "]"
            elif curr_row_num == qbox_header_row:
                print "\r[Key Row " + str(curr_row_num+1) + "]"
            else:
                print "\r[Row " + str(curr_row_num+1) + "]"

            qbox_row_values = worksheet.row_values(curr_row_num)

            for curr_col_num in range(1,6): 

                output = "cell was not read"
                if (worksheet.cell_type(curr_row_num, curr_col_num) == 1):
                    output = unicodedata.normalize('NFKD', worksheet.cell_value(curr_row_num, curr_col_num)).encode('ascii', 'ignore')
                elif (worksheet.cell_type(curr_row_num, curr_col_num) == 0):
                    output = "(Blank)\r" 
                else:
                    output = str(worksheet.cell_value(curr_row_num, curr_col_num)) + "\t" 

                print "[Col " + xlrd.colname(curr_col_num) + "]\t\t" + output
    return 0

# Determines the initial row of each qbox in the worksheet
# Input: an open worksheet
# Output: a list containing the initial row number of each qbox
def locate_qboxHeaderRows(worksheet):
    qboxHeaderRows = []    

    questions_to_omit = ["ID", "", "Active_Positive", "Passive_Positive", "Active_Negative", "Passive_Negative"]

    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1

    curr_row = -1
    while curr_row < num_rows:
        curr_row += 1
        cell_value = worksheet.cell_value(curr_row, 1)
        if cell_value == "Yes":
            if worksheet.cell_value(curr_row, 2).strip() not in questions_to_omit:
                qboxHeaderRows.append(curr_row)

    return qboxHeaderRows

# Determines the ending row of each qbox in the worksheet
# Input: an open worksheet, a list of qbox header row numbers
# Output: a list of tuples containing the number of the first and last row of each qbox
def locate_qbox_footers(worksheet, qboxHeaderRows):
    
    qbox_ending_rows = []

    for qbox_header_row in qboxHeaderRows:
        valid_net_row = qbox_header_row + 2
        while (worksheet.cell_type(valid_net_row, 4) == 2):
            valid_net_row += 1
        qbox_ending_rows.append(valid_net_row)

    qboxDimensions = zip(qboxHeaderRows, qbox_ending_rows)

    return qboxDimensions

# Determines netting rules
# Input: an open worksheet, a list of qbox initial and ending rows
# Output: a list of tuples containing netting rules for each response value
def determine_netting(worksheet, qboxDimensions):
    
    qbox_headers, qbox_footers = zip(*qboxDimensions)

    netting_tuples = []
    for i, qbox_header_row in enumerate(qbox_headers):
        qbox_footer_row = qbox_footers[i]
        curr_row = qbox_header_row + 1
        response_value = []
        netting_category = []
        while curr_row <= qbox_footer_row:
            if (worksheet.cell_type(curr_row, 2) == 2):
                response_value.append(worksheet.cell_value(curr_row, 2))
            else:
                response_value.append("BLANK")
            if (worksheet.cell_type(curr_row, 4) == 2):
                netting_category.append(worksheet.cell_value(curr_row, 4))
            else:
                netting_category.append("BLANK")
            curr_row += 1
        netting_tuples.append(zip(response_value, netting_category))

    return netting_tuples

def printQuestions(questionData):

    print "questions:"

    for question in questionData:
        t_variableName = question[1]
        t_dashboardLabel = question[2]
        t_nettingName = question[6]
        print "    " + str(t_variableName) + ":"
        print "        name: \"" + str(t_dashboardLabel) + "\""
        print "        netted: " + str(t_nettingName) + ""

    print """
    Active_Positive:
        name: Active Positive
        netted: 0
    Passive_Positive:
        name: Passive Positive
        netted: 0
    Active_Negative:
        name: Active Negative
        netted: 0
    Passive_Negative:
        name: Passive Negative
        netted: 0
    passiveActive:
        name: "Total Passive Active"
        type: pasact
"""

def printAnswers(questionData):

    print "answers:"

    for question in questionData:
        t_variableName = question[1]
        t_netNumbers = question[4]
        t_netLabels = question[5]
        print "    " + str(t_variableName) + ":"
        i=0
        while i < len(t_netLabels):
            print "        - name: \"" + str(t_netLabels[i]) + "\""
            print "          code: " + str(float(i+1)) + ""
            i += 1

    print """    passiveActive:
        - name: Active Positive
          code: 1.0
        - name: Active Negative
          code: 2.0
        - name: Passive Positive
          code: 3.0
        - name: Passive Negative
          code: 4.0
        - name: Neutral
          code: 5.0
    Active_Positive: 
        - name: "Very A+ (3 Words) "
          code: 3.0
        - name: "Fairly A+ (2 Words) "
          code: 2.0
        - name: "Not Very A+ (1 Word) "
          code: 1.0
        - name: "Not A+ (0 Words) "
          code: 0.0
    Passive_Positive: 
        - name: "Very P+ (3 Words) "
          code: 3.0
        - name: "Fairly P+ (2 Words) "
          code: 2.0
        - name: "Not Very P+ (1 Word) "
          code: 1.0
        - name: "Not P+ (0 Words) "
          code: 0.0
    Active_Negative: 
        - name: "Very A- (3 Words) "
          code: 3.0
        - name: "Fairly A- (2 Words) "
          code: 2.0
        - name: "Not Very A- (1 Word) "
          code: 1.0
        - name: "Not A- (0 Words) "
          code: 0.0
    Passive_Negative: 
        - name: "Very P- (3 Words) "
          code: 3.0
        - name: "Fairly P- (2 Words) "
          code: 2.0
        - name: "Not Very P- (1 Word) "
          code: 1.0
        - name: "Not P- (0 Words) "
          code: 0.0
"""

def printCustomNetting(questionData):

    netting_to_omit = ["0", "", "TopVersusRestUpTo10", "TopTwoVersusRestUpTo10", "SixToThree"]

    print "customNetting:"

    for question in questionData:
        t_responseValues = question[3]
        t_netNumbers = question[4]
        t_nettingName = question[6]

        if t_nettingName not in netting_to_omit:
            i=0
            print "    " + str(t_nettingName) + ":"
            while i < len(t_responseValues):
                print "        \"" + str(int(t_responseValues[i])) + "\": " + str(int(t_netNumbers[i])) + ""
                i += 1

    print """
    TwoOne:
        "1": 1
        "2": 1
        "3": 2
    0ish:
        "1": 1
        "2": 2
        "3": 3
        "4": 3
        "5": 5
        "6": 6
        "7": 7
    custom2:
        "1": 1
        "2": 2
        "3": 3
        "4": 4
        "5": 5
        "6": 6
        "7": 6
        "8": 6
        "9": 6
        "10": 6
        "11": 6
        "12": 6
        "13": 6
        "14": 6
        "15": 6
        "16": 6
        "17": 6
        "18": 6
        "19": 6
        "20": 6
        "21": 6
    custom3:
        "1": 1
        "2": 1
        "3": 1
        "4": 1
        "5": 2
        "6": 2
    AllToOne:
        "1": 1
        "2": 1
    EightVsFive:
        "1": 1
        "2": 1
        "3": 1
        "4": 1
        "5": 1
        "6": 1
        "7": 1
        "8": 1
        "9": 2
        "10": 2
        "11": 2
        "12": 2
        "13": 2
        "9999": 9999
    Classe_Social:
        "1": 1
        "3": 2
    OneThreeTwo:
        "1": 1
        "2": 2
        "3": 2
        "4": 2
        "5": 3
        "6": 3
    SixToThree:
        "1": 1
        "2": 1
        "3": 2
        "4": 2
        "5": 3
        "6": 3
        "9999": 9999
    q37special:
        "1": 1
        "2": 1
        "3": 2
        "4": 3
        "5": 1
        "6": 4
        "7": 5
        "8": 6
        "9": 7
    top2VersusBottom3:
        "1": 1
        "2": 1
        "3": 2
        "4": 2
        "5": 2
    topVersusRestUpTo5:
        "1": 1
        "2": 2
        "3": 2
        "4": 2
        "5": 2
    TopVersusRestUpTo10:
        "1": 1
        "2": 2
        "3": 2
        "4": 2
        "5": 2
        "6": 2
        "7": 2
        "8": 2
        "9": 2
        "10": 2
        "9999": 9999
        "-99.99": 9999
    TopTwoVersusRestUpTo10:
        "1": 1
        "2": 1
        "3": 2
        "4": 2
        "5": 2
        "6": 2
        "7": 2
        "8": 2
        "9": 2
        "10": 2
        "9999": 9999
        "-99.99": 9999
    bottomVersusRestUpTo4:
        "1": 1
        "2": 1
        "3": 1
        "4": 2
    qKc:
        "1": 1
        "2": 2
        "3": 2
        "4": 2
        "5": 2
        "6": 1
        "7": 2
        "8": 2
        "9": 3
        "9999": 9999
    qKc2:
        "1": 2
        "2": 2
        "3": 2
        "4": 2
        "5": 1
        "6": 2
        "7": 3
        "9999": 9999
    q37:
        "1": 4
        "2": 1
        "3": 3
        "4": 2
        "5": 2
        "6": 2
        "7": 1
        "8": 3
        "9999": 9999
    q53special:
        "1": 1
        "2": 1
        "3": 1
        "4": 2
        "5": 3
        "6": 4
        "7": 1
        "8": 5
        "9": 6
        "10": 6
        "11": 6
        "12": 6
        "13": 6
        "14": 6
        "15": 6
        "16": 6
        "9999": 9999
    Standard5To3:
        "1": 1
        "2": 1
        "3": 2
        "4": 3
        "5": 3
        "9999": 9999
    standard4to2:
        "1": 1
        "2": 1
        "3": 2
        "4": 2
        "9999": 9999
    OneOneTwoThree:
        "1": 1
        "2": 1
        "3": 2
        "4": 3
    OneOneTwo:
        "1": 1
        "2": 2
        "3": 3
        "4": 3
        "9999": 9999
    OneOneTwoTwoTwo:
        "1": 1
        "2": 1
        "3": 2
        "4": 2
        "5": 2
    TwoOneOne:
        "1": 1
        "2": 1
        "3": 2
        "4": 3
    OneOneTwoTwoThree:
        "1": 1
        "2": 1
        "3": 2
        "4": 2
        "5": 3
    RE4:
        "1": 1
        "2": 2
        "3": 3
        "4": 3
        "5": 3
        "6": 3
        "7": 3
        "8": 3
        "9": 3
        "10": 3
    Q14abc:
        "1": 1
        "2": 2
        "3": 3
        "4": 4
        "5": 5
        "6": 5
        "7": 5
AUs:
    - AU02
    - AU12
    - AU04
    - AU09
    - Valence
    - Valence_v2
    - 913
    - Expressiveness
    - Expressiveness_v2
    - Disgust_v3
    - Smile_v2
csvLegend:
    - dummy
"""

def printYamlHeader():

    print '''# the yaml file must define the following attributes:
#
# studies: - a list of studies, each with a magicKey and magicValue String property
#
# variables: - a dict of variables names. Each name maps a var name in CSV to an idx name
#     partipantId - variable name in CSV for participant ID
#     movieId - variable name in CSV for movieId
#
# videos: dict of video names each with name and duration property
#
# questions: dict of questions with name, netted (int), and list of responses
#               each response has a name label and a code. The name is the answer
#               text that will be shown for a given question, and the code
#               is the answer key -- the value that will appear for that
#               answer in the column for that question in the CSV file
# AUs:
#
# options:
#   codeType: int or float
# to use Strings for codes you must put them in quotes as in:
#       code: "3.0a"
#
# MB uses floats for codes
#
studies:
    :
        magicKey: 
        magicValue: live
options:
    codeType: int
    indent: 4
    csv:
        pidIsSessionToken: false
        pidIsNumber: false
    idx:
        returnAllSessions: false
        platform: prod
        pipeline:
            enableViewSequence: true
            viewSequenceValues:
                - "1"
                - "2"
vars:
    csv:
        movieId: MoviNam
        pid: idx
    idx:
        movieId: movieid
        pid: participantId
        viewSequence: viewSequence
videos:
    :
        name: ""
        duration: 
        autoViewSequence:  
'''

def main():

    #######################################################################
    # 1: Input validation. Make sure the arguments and filenames given by  #
    # the user make sense. Make sure the file is a current dashboard spec. #    
    ########################################################################
    
    # Check to see that an argument was given:
    if len(sys.argv) < 2:
        print "usage: ./createSpecification.py <dashboard_spec.xls>"
        sys.exit(0)                                 # TODO: replace with proper exception handling
                                                    # TODO: replace with optparse

    # Verify that the argument is a valid filename:
    dsFilename, dsFileExtension = os.path.splitext(sys.argv[1])
    if not os.path.exists(dsFilename+dsFileExtension):
        print "validation error: you have given the name of a file which does not exist."
        sys.exit(0)                                  #TODO: replace with proper exception handling

    # Verify that the filename points to an xls, xlsx, or xlsm file:
    accepted_formats = ['.xls', '.xlsm', '.xlsx']
    if dsFileExtension.lower() not in accepted_formats:
        print "error: your dashboard spec file must be in one of the following formats: " + \
        '[%s]' % ', '.join(map(str, accepted_formats))
        sys.exit(0)                                  #TODO: replace with proper exception handling

    # Open the workbook:
    workbook = xlrd.open_workbook(dsFilename+dsFileExtension)

    # Verify the workbook contains the expected worksheets:
    expectedWorksheetTitles = ['DashboardSpec-->CS', 'DPSpec-->DP', 'NettingSpec-->idx', \
                                                                   'Net-List', 'Costs', 'Lists']
    unicodeExistingWorksheetTitles = workbook.sheet_names()
    existingWorksheetTitles = map(str, unicodeExistingWorksheetTitles)
    titleIntersection = [filter(lambda x: x in expectedWorksheetTitles, sublist) for sublist in \
                                                                         existingWorksheetTitles]
    titleIntersection = filter(None, titleIntersection)
    if titleIntersection:
        print "validation error: your dashboard specification file has tab names which differ" + \
                                                                         " from what is expected."
        sys.exit(0)                                  #TODO: replace with proper exception handling

    # Verify the workbook is an acceptable version:
    dashboardSpecificationSheet = workbook.sheet_by_name('DashboardSpec-->CS')
    initial_row = dashboardSpecificationSheet.row_values(0)
                                                     #TODO: repair date handling
    date_ok = [a.start() for a in list(re.finditer("Updated 15 March", str(initial_row)))]
    date_ok_2 = [a.start() for a in list(re.finditer("Updated 5 April", str(initial_row)))]
    
    if not (date_ok or date_ok_2):        
        print "validation error: your dashboard spec file is out of date."
        print "This script currently accepts spec files with the dates: 15 March 2013, 5 April 2013"
        sys.exit(0)                                  #TODO: replace with proper exception handling
    else:
        if PRINT_WARNINGS:
            print "file is valid. opening NettingSpec worksheet.\r"

    # Open the NettingSpec worksheet:
    worksheet = workbook.sheet_by_name('NettingSpec-->')

    ###################################################################
    # 2: Import qbox data. Determine which qboxes need to be imported, #
    # their size and shape. Validate these qboxes; import their data.  #
    ####################################################################
    
    # Count qboxes and identify their beginning rows
    qboxHeaderRows = locate_qboxHeaderRows(worksheet)
    
    if PRINT_WARNINGS:
        print str(len(qboxHeaderRows)) + " questions have been found.\r"
    
    # Determine the footer of each qbox, store with header as dimension pairs
    qboxDimensions = locate_qbox_footers(worksheet, qboxHeaderRows)

    if PRINT_DEBUG:
        print_qbox_questions(worksheet, qboxHeaderRows)
        print_qboxes(worksheet, qboxDimensions)

    # Import data from each qbox into a manageable structure
    questionData = import_qboxes(worksheet, qboxDimensions)


    ###################################################################
    # 3: Create yaml file. Output questions, answers, customnetting,   #
    # and header. Validate yaml output.                                #
    ####################################################################

    printYamlHeader()
    printQuestions(questionData)
    printAnswers(questionData)
    printCustomNetting(questionData)

if __name__ == "__main__":
	main()

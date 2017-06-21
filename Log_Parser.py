import glob # Directory Reader module
import xlwt #Excel Writer module
import re # RegEx module
from datetime import datetime
startTime = datetime.now()

#### Creating a workbook
my_xls = xlwt.Workbook(encoding='ascii')

##### Initialising variables to increment
########### TITLES

## Lists comprehension to create iterable variables: add a new tab named after first value
# in the list, then prints every column name in order in a different column

Conversions = ["Site Visits","Ad Views","Orders"]
Cols_name = ["DATE","HOUR","GAID","USER ID","AUDIENCE","EVENT",
"TAG","REFERRER","PAGE NAME","ORDER ID","ORDER VALUE"]

#### Headers for the cols

DATE_col = 0
HOUR_col = 1
GAID_col = 2
CV1_col = 3
AUDIENCE_col = 4
EVENT_col = 5
TAG_col = 6
CV3_col = 7
PAGE_col = 8
ORDER_col = 9
VALUE_col = 10
line_col = 11

### Creating separate sheet with headers per section

my_sheet2 = my_xls.add_sheet("Ad Views")
for col_num, col_name in enumerate(Cols_name):
    my_sheet2.write(0, col_num, col_name)
    print col_num, col_name

 my_sheet3 = my_xls.add_sheet("Orders")
 for col_num, col_name in enumerate(Cols_name):
     my_sheet3.write(0, col_num, col_name)
     print col_num, col_name

my_sheet4 = my_xls.add_sheet("Control")
for col_num, col_name in enumerate(Cols_name):
    my_sheet4.write(0, col_num, col_name)
    print col_num, col_name


##### RegEx variables to match
regexDATE = r"\d{2}\/\w{3}/\d{4}" # matches xx/xx/xxxx
regexHOUR = r"\:(\d{2}):(\d{2}):(\d{2})"
regexGAID = r"\gaid=(\w*.\w*)"
regexCV1 = r"CV1=(\w*)"
regexAUDIENCE = r"AudienceID=(\w*)%(\d\w*)%(\d\w*)|AudienceID=(\w*)%(\d\w*)|AudienceID=(\w*)"
regexEVENT = r"EventType=(\w*)"
regexTAG = r"TagID=(\w*)"
regexCV3 = r"CV3=(\w*)[:][\/][\/](\w*)[.](\w*)[.](\w*)[\/](\w*)"
regexPAGE = r"PageName=\/(\w*)\.(\w*)"
regexORDER = r"CV9=(\w*)"
regexVALUE = r"CV10=(\w*.\w*)"

#################################### USER INTERFACE
########################################################
print ""
print ""
Day = (raw_input('Please input the date of the day (2 digits: eg 01,30) you\'re interested in: '))
print ""
Month = (raw_input('Please input the month (3 char: eg Jan, Nov) you\'re interested in: '))
print ""
print ""

############################################# GENERAL VARIABLES
folder = 'C:\\Users\\Data\\Desktop\\Logs\\' + Month + '\\' + Day + '/*.log'
# file1 = 'C:\\Users\\Data\\Desktop\\Logs\\' + Month + '\\' + Day + '/????????????????????????.184.log'
# file2 = 'C:\\Users\\Data\\Desktop\\Logs\\' + Month + '\\' + Day + '/????????????????????????.196.log'
# file3 = 'C:\\Users\\Data\\Desktop\\Logs\\' + Month + '\\' + Day + '/????????????????????????.253.log'
xls_saved = 'Naked_Activity_' + Day + '_' + Month +'.xls'

######################################## All Site ###############################################
 row_Visit = 1

 Visits = ['SiteVisit']

 for file in glob.iglob(folder): #iterate through files in the directory
     substr = "TagID" # Substring to look for
     with open (file, 'rt') as in_file:
         for line in in_file:
             if any(s in line for s in Visits): # Match Parameter names
                 index = 0
                 str = line
                 while index < len(str):
                     index = str.find(substr, index)
                     # if substring isn't found in the line, stop searching and go to next instructions (restart process on next log entry)
                     if index == -1:
                         break
                     # If substring found in the log entry: create variable storing the results of re.findall functions, print first match in column1
                     matchesDATE = re.findall(regexDATE,str) # Find all strings matching regular expression 1 (regex1)
                     matchesHOUR = re.findall(regexHOUR,str) # Find all strings matching regular expression 2 (regex2)
                     matchesGAID = re.findall(regexGAID,str) # Find all strings matching regular expression 3 (regex3)
                     matchesCV1 = re.findall(regexCV1,str)# Find all strings matching regular expression 4 (regex4)
                     matchesAUDIENCE = re.findall(regexAUDIENCE,str) # Find all strings matching regular expression 5 (regex5)
                     matchesEVENT = re.findall(regexEVENT,str) # Find all strings matching regular expression 6 (regex6)
                     matchesTAG = re.findall(regexTAG,str) # Find all strings matching regular expression 7 (regex7)
                     matchesCV3 = re.findall(regexCV3,str) # Find all strings matching regular expression 8 (regex8)
                     matchesPAGE = re.findall(regexPAGE,str) # Find all strings matching regular expression 9 (regex9)
                     # Print chunks of the log entry that match each regex in its own separate column
                     for match1 in matchesDATE: # DATE
                         print match1
                         my_sheet1.write(row_Visit,DATE_col,match1)

                     for match2 in matchesHOUR: # HOUR
                         print match2
                         my_sheet1.write(row_Visit,HOUR_col,match2)

                     for match3 in matchesGAID: # GAID
                         print match3
                         my_sheet1.write(row_Visit,GAID_col,match3)

                     for match4 in matchesCV1:# CV1
                         print match4
                         my_sheet1.write(row_Visit,CV1_col,match4)
                         break

                     for match5 in matchesAUDIENCE: # AUDIENCE
                         print match5
                         my_sheet1.write(row_Visit,AUDIENCE_col,match5)
                         break

                     for match6 in matchesEVENT: # EVENT
                         print match6
                         my_sheet1.write(row_Visit,EVENT_col,match6)
                         break

                     for match7 in matchesTAG: # TAG
                         print match7
                         my_sheet1.write(row_Visit,TAG_col,match7)
                         break

                     for match8 in matchesCV3: # CV3
                         print match8
                         my_sheet1.write(row_Visit,CV3_col,match8)
                         break

                     for match9 in matchesPAGE: # PAGE
                         print match9
                         my_sheet1.write(row_Visit,PAGE_col,match9)
                         break

                     row_Visit += 1
                     index += len(str)

####################################### Ad Views ###############################################
print "Starting Ad Views Script"

startTimeAdView = datetime.now()

row_Ad = 1
Audiences = ['GET /Logger?','EventType'] # List of parameters compounded with the substring

for file in glob.iglob(folder): #iterate through files in the directory
    substr = "ClientID" # Substring to look for
    with open (file, 'rt') as in_file:
        for line in in_file:
            if all(s in line for s in Audiences): # Changed any to all to match logger as well
                str = line
                index = 0
                while index < len(str):
                    index = str.find(substr, index)
                    # if substring isn't found in the line, stop searching and go to next instructions (restart process on next log entry)
                    if index == -1:
                        break
                    else:
                        my_sheet2.write(row_Ad, line_col,line)
                        # If substring found in the log entry: create variable storing the results of re.findall functions, print first match in column1
                        matchesDATE = re.findall(regexDATE,str) # Find all strings matching regular expression 1 (regex1)
                        matchesHOUR = re.findall(regexHOUR,str) # Find all strings matching regular expression 2 (regex2)
                        matchesGAID = re.findall(regexGAID,str) # Find all strings matching regular expression 3 (regex3)
                        matchesAUDIENCE = re.findall(regexAUDIENCE,str) # Find all strings matching regular expression 5 (regex5)
                        # Print chunks of the log entry that match each regex in its own separate column
                        for match1 in matchesDATE: # DATE
                            my_sheet2.write(row_Ad,DATE_col,match1)
                            break

                        for match2 in matchesHOUR: # HOUR
                            my_sheet2.write(row_Ad,HOUR_col,match2)
                            break

                        for match3 in matchesGAID: # GAID
                            my_sheet2.write(row_Ad,GAID_col,match3)
                            break

                        for match5 in matchesAUDIENCE: # AUDIENCE
                            my_sheet2.write(row_Ad,AUDIENCE_col,match5)
                            break

                        row_Ad += 1
                        index += len(str)

print datetime.now() - startTimeAdView

####################################### Control ###############################################
print "Starting Control Script"

startTimeControl = datetime.now()

Control = ['GET /Logger?','EventType'] # List of parameters compounded with the substring
row_Control = 1

for file in glob.iglob(folder): #iterate through files in the directory
    substr = "ClientID" # Substring to look for
    with open (file, 'rt') as in_file:
        for line in in_file:
            if all(s in line for s in Control): # Changed any to all to match logger as well
                str = line
                index = 0
                while index < len(str):
                    index = str.find(substr, index)
                    # if substring isn't found in the line, stop searching and go to next instructions (restart process on next log entry)
                    if index == -1:
                        break
                    else:
                        my_sheet4.write(row_Control, line_col,line)
                        # If substring found in the log entry: create variable storing the results of re.findall functions, print first match in column1
                        matchesDATE = re.findall(regexDATE,str) # Find all strings matching regular expression 1 (regex1)
                        matchesHOUR = re.findall(regexHOUR,str) # Find all strings matching regular expression 2 (regex2)
                        matchesGAID = re.findall(regexGAID,str) # Find all strings matching regular expression 3 (regex3)
                        matchesAUDIENCE = re.findall(regexAUDIENCE,str) # Find all strings matching regular expression 5 (regex5)
                        # Print chunks of the log entry that match each regex in its own separate column
                        for match1 in matchesDATE: # DATE
                            my_sheet4.write(row_Control,DATE_col,match1)
                            break

                        for match2 in matchesHOUR: # HOUR
                            my_sheet4.write(row_Control,HOUR_col,match2)
                            break

                        for match3 in matchesGAID: # GAID
                            my_sheet4.write(row_Control,GAID_col,match3)
                            break

                        for match5 in matchesAUDIENCE: # AUDIENCE
                            my_sheet4.write(row_Control,AUDIENCE_col,match5)
                            break

                    row_Control += 1
                    index += len(str)

print datetime.now() - startTimeControl




print datetime.now() - startTime
my_xls.save(xls_saved)

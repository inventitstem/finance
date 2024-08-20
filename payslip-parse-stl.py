import pymupdf
import os
import pandas as pd
import pkg_resources
from pprint import pprint

def add_dict(key,value,dictionary,line):
    #print("Hello from a function")
    if key in dictionary:
        print("Key:" + key + " already exists")
        key = key+"1"
    dictionary[key] = value
    print(str(line) + ":"+ key + " = " + value)
    
def print_env():
    # requirements.txt is generated with pipfreeze
    env_list=[]
    print("Getting installed packages")
    installed_packages = pkg_resources.working_set
    #print(str(len(installed_packages)))
    installed_packages_list = sorted(["%s==%s" % (i.key, i.version) for i in installed_packages])
    for package in installed_packages_list:
        print(package)
    return(installed_packages_list)
    

mode=[]
# Format of the payslip is as follows
#- Payslip
earningColumns=["QUANTITY","RATE"]
totalColumns=["THIS PAY","YTD"]

mode=[{"name": "FILE", "tag": "zzz","concatenate":True,"header":False,"useHeader":False,"columns":"","fields":1}, #-- Recipient ["None"]
    {"name": "RECIPIENT", "tag": "zzz","concatenate":True,"header":False,"useHeader":False,"columns":"","fields":1000}, #-- Recipient ["None"]
    {"name": "PAID BY", "tag": "paid by","concatenate":True,"header":True,"useHeader":False,"columns":"","fields":1000},#-- Paid By -> "PAID BY" to "EMPLOYMENT DETAILS"
    {"name": "EMPLOYMENT DETAILS", "tag": "employment details","concatenate":False,"header":True,"useHeader":True,"columns":"","fields":1000},#-- Employment Details -> "EMPLOYMENT DETAILS" to "Pay Period"
    {"name": "PAY PERIOD", "tag": "pay period","concatenate":False,"header":False,"useHeader":True,"columns":"","fields":1000},#-- Pay Period -> "Pay Period" -> "Payment Date"
    {"name": "PAYMENT SUMMARY", "tag": "payment date","concatenate":False,"header":False,"useHeader":True,"columns":"","fields":1000},#-- Payment Summary -> "Payment Date" -> "THIS PAY": Single Row
    {"name": "TABLE COLUMNS", "tag": "this pay","concatenate":False,"header":False,"useHeader":False,"columns":"","fields":1000},#-- Table Columns -> "THIS PAY" -> "EARNINGS": First two rows give the column headings i.e. THIS PAY and YTD
    {"name": "EARNINGS", "tag": "earnings","concatenate":False,"header":True,"useHeader":False,"columns":earningColumns + totalColumns,"fields":1000},#-- Earnings -> "EARNINGS" -> "DEDUCTIONS": First to rows after "EARNINGS" give the column headings for this section. In this section it is Quantity, Rate, THIS PAY, YTD
    {"name": "DEDUCTIONS", "tag": "deductions","concatenate":False,"header":True,"useHeader":False,"columns":totalColumns,"fields":1000},#-- Deductions -> "Deductions" -> "PAYMENT Details"
    {"name": "PAYMENT DETAILS", "tag": "payment details","concatenate":False,"header":False,"useHeader":False,"columns":"","fields":1000},#-- Payment Details -> "Payment Details" -> "Employer Contributions"
    {"name": "EMPLOYER CONTRIBUTIONS", "tag": "employer contributions","concatenate":False,"header":True,"useHeader":False,"columns":totalColumns,"fields":1000},#-- Employer Contributions -> "Employer Contributions" -> "TOTAL" + 2 rows
    {"name": "OTHER", "tag": "zzz","concatenate":False,"header":False,"useHeader":False,"columns":"","fields":1000},#-- Other -> "Total" + 2 rows -> EOD
    {"name": "END", "tag": "zzz","concatenate":False,"header":False,"useHeader":False,"columns":"","fields":1000}#-- Other -> "Total" + 2 rows -> EOD
]


validEntries=["TOTAL"]

startDate= pd.to_datetime('09/08/2021',format='%d/%m/%Y')
endDate=pd.to_datetime('22/08/2024',format='%d/%m/%Y')
#startDate= pd.to_datetime('09/02/2024',format='%d/%m/%Y')
#endDate=pd.to_datetime('22/03/2024',format='%d/%m/%Y')

startDateString=str(startDate.month) + " " + str(startDate.year)
endDateString=str(endDate.month) + " " + str(endDate.year)

#Create dates
dateList=[]
dateValue= startDate
while dateValue <= endDate:
    dateList.append(dateValue.month_name() + " " + str(dateValue.year))
    dateValue=dateValue+pd.DateOffset(months=1)
    
print("Found " + str(len(dateList)) + "months employment")

#Get all files in the directory
path="C://Users//RichardGray//OneDrive//STL//Payslip"

csv_file_path = 'summary.csv'
csv_file_path_unique = 'summary_unique.csv'
result = pd.DataFrame()
    
# to store files in a list
files = []

# dirs=directories

for (root, dirs, file) in os.walk(path):
    for f in file:
        if '.pdf' in f:
            files.append(os.path.join(root, f))
            
print ("Analysing " + str(len(files)) + " files")
       
        
for payslip in dateList:
    print("Analysing " + payslip)
    
    #Find Payslip
    matches = [match for match in files if payslip in match]
    
    print("Found " + str(len(matches)) +" files")
    
    for i in range(len(matches)):
        if "P60" in matches[i]:
            print("P60 - Ignore")
            break
        else:
            doc = pymupdf.open(matches[i]) # open document
            numPages=len(doc)
            print(str(numPages) + "pages found")

            for p in range(numPages):
                print("Analysing page "+ str(p))
                page = doc[p] # get the 1st page of the document
                text = matches[i] + "\n" + page.get_text()#.encode("utf8") # get plain text (is in UTF-8)
                if payslip == "March 2024":
                    print(text)
                data = text.split("\n")
                df=pd.DataFrame(data,columns=[payslip])
                #Extract the information from the data
                thisdict={"filename": matches[i]}
                modeNum=0
                newModeNum=0
                dataEntry=""
                fieldCount=mode[modeNum]["fields"]
                skip=0
                for j in range(len(data)):
                    header=False
                    #if newModeNum != modeNum:
                        #dataEntry=""
                        #modeNum = newModeNum
                        #fieldCount = mode[modeNum+1]["fields"] + j
                    if skip <= j:
                        print(str(j)+'->Mode:'+str(modeNum)+ '. Checking "' + data[j] + '" for "' + mode[modeNum+1]["tag"] + '" or line ' + str(fieldCount))
                        # Search for next mode
                        if modeNum == len(mode)-2:
                            break
                        if mode[modeNum+1]["tag"].lower() in data[j].lower() or j >= fieldCount:
                            print("")
                            #Handle current data
                            
                            #Update any concatenated fields
                            if mode[modeNum]["concatenate"]==True:
                                print(dataEntry)
                                thisdict[mode[modeNum]["name"]]=dataEntry
                                add_dict(mode[modeNum]["name"],dataEntry,thisdict,j)
        
                            # Update mode number
                            #newModeNum = modeNum+1
                            modeNum = modeNum+1
                            fieldCount = mode[modeNum+1]["fields"] + j
                            dataEntry=""
                            print("Detected mode " + str(modeNum) + ":" + mode[modeNum]["name"] + " on line " + str(j))
                            if mode[modeNum]["header"] == True:
                                header=True
                            
                        
                        #Perform action on current mode
                        if mode[modeNum]["concatenate"] == True and header == False:
                            if dataEntry == "":
                                dataEntry = data[j]
                            else:
                                dataEntry= dataEntry + ", " + data[j]
                                
                        elif mode[modeNum]["useHeader"] == True and header == False:
                            #print(data[j].split(': '))
                            #thisdict[data[j].split(":")[0]]=data[j].split(":")[1].strip()
                            add_dict(data[j].split(":")[0],data[j].split(":")[1].strip(),thisdict,j)
                        elif mode[modeNum]["columns"] != "" and header == False:
                            #Ignore column headings
                            if data[j] == "TOTAL":
                                #thisdict[mode[modeNum]["name"] + "_" + data[j] + "_" + totalColumns[0]] = data[j+1]
                                #print(mode[modeNum]["name"] + "_" + data[j] + "_" + totalColumns[0] + " = " + data[j+1])
                                add_dict(mode[modeNum]["name"] + "_" + data[j] + "_" + totalColumns[0],data[j+1],thisdict,j+1)
                                #thisdict[mode[modeNum]["name"] + "_" + data[j] + "_" + totalColumns[1]] = data[j+2]
                                #print(mode[modeNum]["name"] + "_" + data[j] + "_" + totalColumns[1] + " = " + data[j+2])
                                add_dict(mode[modeNum]["name"] + "_" + data[j] + "_" + totalColumns[1],data[j+2],thisdict,j+2)
                                len(mode[modeNum]["columns"])
                                skip=j+3                            
                            elif data[j] not in earningColumns and not data[j].startswith("£") and data[j] !=" ":
                                #print('"' + data[j] + '" is valid')
                                if payslip == "August 2021" and j == 28:
                                    columnValues=["THIS PAY","YTD"]
                                elif payslip == "August 2021" and j == 31:
                                    columnValues=["QUANTITY","RATE"]
                                elif (payslip == "February 2024" or payslip == "March 2024") and j == 32:
                                    columnValues=["YTD"]
                                else:
                                    columnValues=mode[modeNum]["columns"]
                                for k in range(len(columnValues)):
                                    print(k)
                                    #thisdict[mode[modeNum]["name"] + "_" + data[j] + "_" + mode[modeNum]["columns"][k]] = data[j+k+1]
                                    #print(str(j+k+1) + ":"+ mode[modeNum]["name"] + "_" + data[j] + "_" + mode[modeNum]["columns"][k] + " = " + data[j+k+1])
                                    add_dict(mode[modeNum]["name"] + "_" + data[j] + "_" + columnValues[k],data[j+k+1],thisdict,j+k+1)
                                skip=j+len(columnValues)+1
                        if mode[modeNum]["name"] == "PAYMENT DETAILS":
                            if data[j] == "Electronic Transfer":
                                #thisdict[mode[modeNum]["name"] + "_" + data[j]] = data[j+4]
                                #print(mode[modeNum]["name"] + "_" + data[j] + " = " + data[j+4])
                                add_dict(mode[modeNum]["name"] + "_" + data[j],data[j+4],thisdict,j+4)
                                
                        if modeNum == 10 and data[j] == "TOTAL":
                            fieldCount=j+2
                        

            #print dictionary
            print(thisdict)
            #convert into panda data frame
            df_dict = pd.Series(thisdict, name= payslip)
            #df_dict.index.name = 'Date'
            #df_dict.reset_index()
            df_dict.to_csv(csv_file_path_unique, index=True,encoding="utf-8-sig")
            
            #result.reset_index(inplace=True)
            result = pd.concat([result,df_dict],axis=1)
                
                #tabs = page.find_tables() # locate and extract any tables on page
                #numTables = len(tabs.tables)
                #print(f"{len(tabs.tables)} found on {page}") # display number of found tables

                #if tabs.tables:  # at least one table found?
                #    for t in range(numTables):
                #        print("Analysing table "+ str(t))
                        #pprint(tabs[t].extract())  # print content of first table

result.to_csv(csv_file_path, index=True,encoding="utf-8-sig")
# Step 4: Write the DataFrame back to the CSV file

print_env()

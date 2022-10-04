import pandas as pd, time, string, os, pickle, xlsxwriter, openpyxl
import DictionPickle as dp


#Pull in dictionary

#Read the words from input
def RecordWords():
#Create variables
    global words
    global create_list
    words = input()

#Punctuation replaced by spacing
    punct_1 = ['-','.',',','—',';',':','?','!','/','(',')']
    punct_2 = ['"']
    for p in punct_1:
        words = words.replace(p,' ')
    for p in punct_2:
        words = words.replace(p,'')
        
#Use split to create a list from string
    words = string.capwords(words)
    create_list = words.split()
    
#Remove dupes from the list
    word_list = [*set(create_list)]
    
#Add entries to the current dictionary
#We will use the current dictionary to make a new dictionary
    for w in word_list:
        if len(curr_dict) != 0:
            wl_last_entry = list(curr_dict)[-1]
        else:
            wl_last_entry = 0
                
        wl_new_entry = wl_last_entry + 1
        curr_dict[wl_new_entry] = w
    
#New dictionary
    
#Variables
    global new_dict
    new_list = []
    new_dict = {}
    num = 0
 
#Separate current dictionary values, remove dupes, and alphabetize in a new list 
    curr_dict_val = curr_dict.values()
    for v in curr_dict_val:
        new_list.append(v)
    new_list = [*set(new_list)]
    new_list = sorted(new_list)
    
#Add the new list values to a new dictionary 
    for i in new_list:
        num = num + 1
        new_dict[num] = i
    #print(new_dict)
    
#Check both new and old dicts, prompt user if they would like to update the current dictionary
    if new_dict != old_dict:
        print('You have new words to be added to your dictionary. Would you like to add them now?')
        AddToDict()
    else:
        print('You did not use any new words this time.')
        print(new_dict)
        print(old_dict)
 
#Check if user wants to update the current dictionary 
def AddToDict():
    yes_no_in = input()
    yes_no = yes_no_in.lower()
    if yes_no == 'yes':

#If yes, current dictionary is updated to be the new dictionary
        curr_dict = new_dict
        dp.save_diction(curr_dict)
        print('Dictionary updated successfully.')
        print(curr_dict)
        time.sleep(1)
        print('. . .')
        time.sleep(2)
        print('Would you like to export the current dictionary?')
        exportCheck()
        
#If no, current dictionary is kept as old dictionary
    elif yes_no == 'no':
        curr_dict = old_dict
        print('No updates were made.')
        print(curr_dict)
        time.sleep(1)
        print('. . .')
        time.sleep(2)
        print('Would you like to export the current dictionary?')
        exportCheck()
    else:
        print("Sorry, I didn't get that.")
        AddTryAgain() 
    
def AddTryAgain():
    time.sleep(2)
    print("Would you like to add new words to your dictionary? (Type 'yes' or 'no')")
    AddToDict()
    
#Check if user wants to export        
def exportCheck():
    yes_no_in = input()
    yes_no = yes_no_in.lower()
    if yes_no == 'yes':
    
    #Check if the file was created this session
        if newFile == "True":
            dp.UpdateFile("filename.pickle")
            time.sleep(1)
            print("Exporting...")
            time.sleep(1)
            ExportDict()
            
    #If the file is existing, ask if the user wants to Replace or Save new
        else:
            print("Would you like to Replace your current file or Save a new version? (Type 'Replace' or 'Save new')")
            SaveOrRep()
    elif yes_no == 'no':
        print('Dictionary was not exported.')
    else:
        print("Sorry, I didn't get that.")
        ExTryAgain()
        
def ExTryAgain():
    time.sleep(2)
    print("Would you like to add new words to export your dictionary? (Type 'yes' or 'no')")
    
#Check user input for Replace or Save new
    
def SaveOrRep():
    SavRep_in = input()
    SavRep = SavRep_in.lower()
    
    if SavRep == 'replace':
        print("File will be replaced using existing filename: " + fn.param)
        time.sleep(2)
        print("Exporting...")
        time.sleep(1)
        ExportDict()
    elif SavRep == 'save new':
        dp.UpdateFile("filename.pickle")
        fn = dp.load_object("filename.pickle")
        print("File will be saved as a new version: " + fn.param)
        time.sleep(2)
        print("Exporting...")
        time.sleep(1)
        ExportDict()
    else:
        print("I'm sorry, I didn't get that.")
        time.sleep(1)
        savTryAgain()
    
def savTryAgain():
    print("Do you want to Replace your current file or save a new version? (Type 'Replace' or 'Save new')")
    SaveOrRep()
    
def ExportDict():
    
#Set variables
    obj = dp.load_object("filename.pickle")
    filename = obj.param
    data = new_dict
    
    data_val = list(data.values())

    df = pd.DataFrame({"Number" : data.keys(),
                        "Term" : data_val})
                        
    dictTable = (df.T)
    
    print(dictTable)
    
    header_names = ['Number','Term']
#Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    
#Convert the dataframe to an XlsxWriter Excel object. Note that we turn off
#the default header and skip one row to allow us to insert a user defined
#header.

    df.to_excel(writer, sheet_name = 'Sheet1', startrow = 1, index = False, header = False)
    
#Get the xlsxwriter workbook and worksheet objects.

    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
#Add a header format
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1})
        
#Write the column headers with the defined format.

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
#Close the Pandas Excel writer and output the Excel file.
    writer.save()
    print("Excel saved successfully.")
        
#Ask user to enter in text
def PromptUser():
#Load the filename in case it was just created
    global fn
    fn = dp.load_object("filename.pickle")
    
#Prompt user to enter text
    print("Please enter in some text.")
    RecordWords()
    
def checkFilename():
        
#If the filename already exists, load it
    fn = dp.load_object("filename.pickle")
    print("You have an existing file: " + fn.param)
    time.sleep(1)
    print("Would you like to use your existing filename?")
    makeNew()
        
def makeNew():
#Create a variable to detect if a new file is created this session   
    global newFile
    
    New_YN = input()
    New_YN = New_YN.lower()
    
    if New_YN == 'no':
        print("Would you like to create a new file or load an existing one? (Type 'Create New' or 'Load File')")
        NeworLoad()
      
    elif New_YN == 'yes':
        newFile = "False"
        old_dict = dict(loadDict)
        curr_dict = dict(loadDict)
        
        print("Existing file will be used.")
        time.sleep(1)
        PromptUser()
        
    else:
        print("Sorry, I didn't get that.")
        time.sleep(2)
        makeNewAgain()
        
def makeNewAgain():
    print("Would you like to use your existing filename? (Type yes or no)")
    makeNew()
    
#Check if user wants to create a new file or load an existing one

def NeworLoad():
    NL_in = input()
    NL = NL_in.lower()
    
#Create a variable to detect if a new file is created this session   
    global newFile
    
    if NL == 'create new':
#Reset dictionary variables
        old_dict = {}
        curr_dict = {}
        
#Update variable that is tracking new file       
        newFile = "True"
        
#Reset the pickled dictionary
        dp.save_diction(curr_dict)

#Prompt user to create a new filename
        dp.fileprompt()
        
#Continue to next step
        time.sleep(1)
        PromptUser()
        
    elif NL == 'load file':
        newFile = "False"
        print("Please enter in the filename (without .xlsx).")
        getFile()
        
    else:
        print("Sorry, I didn't get that.")
        NLAgain()
        
def NLAgain():
    
    print("Would you like to create a new file or load an existing one? (Type 'Create New' or 'Load File')")
    NeworLoad()
    
def getFile():
    global old_dict
    global curr_dict

#Get the filename and store it
    dp.createFile()
    loadFN = dp.filename
    print("Your filename is: " + loadFN)
    time.sleep(1)
    print("File stored.")
    time.sleep(1)
    
 #Grab workbook sheets
    loadData = pd.ExcelFile(loadFN)
    
#Create dataframe
    df = loadData.parse('Sheet1')
    
#Load workbook, grab Sheet1
    ps = openpyxl.load_workbook(loadFN)
    sheet = ps['Sheet1']
    
#Create a temporary dictionary 
    temp_dict = {}
    
#Use a for loop to pull out values from the sheet
    for r in range(2, sheet.max_row + 1):
        number = sheet['A' + str(r)].value
        term = sheet['B' + str(r)].value
        
    #Put values in the dictionary
        temp_dict[number] = term
    
#Save temp dictionary
    dp.save_diction(temp_dict)
    
#Load dictionary and set it to old and current dicts
    loadDict = dp.load_object("dictionary.pickle")
    old_dict = dict(loadDict)
    curr_dict = dict(loadDict)
    PromptUser()
    

#Check for existing dictionary

def dictCheck():
    global old_dict
    global curr_dict
    global loadDict
    
    old_dict = {}
    curr_dict = {}
    
    if os.path.isfile("dictionary.pickle"):
        loadDict = dp.load_object("dictionary.pickle")
        print("Loaded most recent dictionary...")
        time.sleep(1)
        checkFilename()
        
    else:
        print("No dictionary found.")
        time.sleep(1)
        print("Would you like to create a new file or load an existing one? Type 'Create New' or 'Load File')")
        NeworLoad()

dictCheck()
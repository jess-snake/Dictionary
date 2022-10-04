import pickle, re 
 
class SaveFile():
    def __init__(self, param):
        self.param = param
        
#Save the variable
 
def save_object(obj):
    try:
        with open("filename.pickle", "wb") as f:
            pickle.dump(obj, f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception as ex:
        print("Error during pickling object (Possibly unsupported):", ex)
        
    #Show the variable
    
    #obj = load_object("filename.pickle")
    #print(obj.param)
    
    #Because obj is also a param, calling it here will show once,
    #while the save_object and load_object functions technically process twice
    #due to the save_object function creating the file and updating the file.
    
#Load the variable
    
def load_object(filename):
    try:
        with open(filename, "rb") as f:
            return pickle.load(f)
    except Exception as ex:
        print("Error during unpickling object (Possibly unsupported):", ex)
    
#Ask user for variable input

def fileprompt():
    print("Please enter a name for your file.")
    createFile()
    
def createFile():
    global text
    text = input()
    global filename
    filename = text + ".xlsx"
    
#Use class and save_object function to store input
    
    obj = SaveFile(filename)
    save_object(obj)
    
#Update file version number
    
def UpdateFile(filename):

#Assumes file can be accessed so it's only called once
    loadFile = pickle.load(open(filename, "rb"))
    
    if loadFile:
        fn = loadFile.param
        
        verCheck = re.search('_V(.+?).xlsx', fn)
        
        if verCheck:
        #Break down original filename to isolate num.xlsx 
            breakdown_1 = fn.partition('_V')
            grabIt_1 = breakdown_1[2]
            grabIt_1 = str(grabIt_1)
            
        #Break down to isolate the num only
            breakdown_2 = grabIt_1.partition('.xlsx')
            grabIt_2 = breakdown_2[0]
            grabIt_2 = int(grabIt_2)
        
        #Reassign version number
            versionNum = grabIt_2 + 1
            versionNum = str(versionNum)
            
        #Grab 1st portion of filename
            FN_part1 = breakdown_1[0]
            FN_part1 = str(FN_part1)
            
        #Create new filename
            fnNew = FN_part1 + "_V" + versionNum + ".xlsx"
            
        else:
        
            #print("File version will be saved as 'V1'.")

            versionNum = 1

            versionNum = str(versionNum)
            
            #Break down original filename to 1st portion
            breakdown_1 = fn.partition('.xlsx')
            FN_part1 = breakdown_1[0]
            FN_part1 = str(FN_part1)

            #Create new filename
            fnNew = FN_part1 + "_V" + versionNum + ".xlsx"
            
        obj = SaveFile(fnNew)
        save_object(obj)
        
    else:
        print("Couldn't load file.")
        
#Save the dictionary
 
def save_diction(obj):
    try:
        with open("dictionary.pickle", "wb") as f:
            pickle.dump(obj, f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception as ex:
        print("Error during pickling object (Possibly unsupported):", ex)
     
    #obj = load_object("dictionary.pickle")
    #print(obj)

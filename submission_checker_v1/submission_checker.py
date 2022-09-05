import os, zipfile, csv, shutil

#############################################################################################
#Gets number of questions from csv using Student ID.
#############################################################################################
def get_tasks(HW_no):
    #open file with reader
    with open(os.path.join(program_dir,'homeworks.csv')) as File:
        reader = csv.reader(File, delimiter=',', quotechar=',',quoting=csv.QUOTE_MINIMAL)
        #row counter to skip header row
        rowNr = 0
        #loop through rows to find id
        for row in reader:
            if rowNr >= 1:
                
                if int(row[0])==HW_no:
                    return row[1:]
            rowNr= rowNr+1

#############################################################################################
#Gets author name from csv using Student ID.
#############################################################################################
def get_author(id):
    #open file with reader
    with open(os.path.join(program_dir,'students.csv')) as File:
        reader = csv.reader(File, delimiter=',', quotechar=',',quoting=csv.QUOTE_MINIMAL)
        #row counter to skip header row
        rowNr = 0
        #loop through rows to find id
        for row in reader:
            if rowNr >= 1:
                
                if int(row[0])==id:
                    return row[1]
            rowNr= rowNr+1
        
        return "NOT FOUND"

#############################################################################################
#Extracts zip files in working directory.
#############################################################################################
def zip_extract():
    extension = ".zip"
    # change directory from working dir to dir with files
    os.chdir(root_dir) 
    print("Extracting...\n")
    for root, dirs, files in os.walk(root_dir): # loop through items in dir
        for file in files:
            if file.endswith(extension): # check for ".zip" extension
                file_name = os.path.basename(file) # get full path of files
                file_path = os.path.join(root, file)
                #print(file_name)   #debug
                zip_ref = zipfile.ZipFile(file_path) # create zipfile object
                folder_name = os.path.join("extracted",file_name.split(".")[0])
                try:
                    os.makedirs(folder_name) #create directory
                    print(folder_name+" created.")  #debug
                except:
                    print(folder_name+" already exists!")

                zip_ref.extractall(folder_name) # extract file to dir
                zip_ref.close() # close file
                os.remove(file_path)
            
    print("\nExtracting complete.\n")
    
    for folder in  os.listdir(root_dir):
        if os.path.isdir(folder) and folder != "extracted" and folder != "sc_files":
            try:
                shutil.rmtree(folder)
            except:
                print("No such folder")
    

#############################################################################################
#Checks file for correct Student-ID and author name.
#############################################################################################    
def check_file(file):
    #open file
    c= open(os.path.join(dir_path,file),"r",encoding='cp437')
    
    #loop through lines
    found_auth=0
    found_id=0
    for line in c:
        #print(line)
        if line.find("@author") != -1:
            file_auth=line[line.find("@author")+8:-1]
            found_auth=1
            #check if author entry exists
            global author
            if author == "NOT FOUND":
                if file_auth.lower().replace(" ","") !="" and file_auth.lower().replace(" ","") != "Matthias Krauledat".lower().replace(" ",""):
                    write_student([id,file_auth])
                    
                    author = file_auth
                    t_author[t_no-1]=1
            else:
                if ''.join(filter(str.isalpha, file_auth)).lower().replace(" ","")==''.join(filter(str.isalpha, author)).lower().replace(" ",""):
                    #set author flag
                    t_author[t_no-1]=1
                
        if line.find("@id") != -1:
            file_id=line[line.find("@id")+4:-1]
            found_id=1
            global errors
            try:
                if int(''.join(filter(str.isdigit, file_id)))==id:
                    #set id flag
                    t_id[t_no-1]=1
            except:
                errors = errors +"--ID error: "+dir_path+"\\"+file+"\n"
        
        if found_auth==1 and found_id==1:
            break

#############################################################################################
#Export data to .csv file.
#############################################################################################    
def export_data():
    os.chdir(program_dir)
    with open('Report.csv', 'w',newline='') as csvfile:
        filewriter = csv.writer(csvfile, delimiter=',',quotechar='|', quoting=csv.QUOTE_MINIMAL)
        headers=['ID', 'Name','Ex No','Tasks']
        for i in range(1, tasks+1):
            headers = headers + ['T'+str(i)+' P'] +['T'+str(i)+' Au']+['T'+str(i)+' ID']+['T'+str(i)+' Grade']
        filewriter.writerow(headers)
        # for i in data:
            # filewriter.writerow(data[i])
        filewriter.writerows(data)
        print("\nExport complete!\n")
 
#############################################################################################
#Export data to .csv file.
#############################################################################################    
def write_student(data):
    os.chdir(program_dir)
    with open('students.csv', 'a+',newline='') as csvfile:
        filewriter = csv.writer(csvfile, delimiter=',',quotechar='|', quoting=csv.QUOTE_MINIMAL)
        filewriter.writerow(data)
 
#############################################################################################
#Main program flow.
#############################################################################################
print("\n____________________________\nSUBMISSION CHECKER\n____________________________\n\n")

errors="\n_________errors_________\n"
root_dir = os.getcwd() #get current directory
ext_dir = root_dir+"\\extracted"
program_dir = root_dir+"\\sc_Files"
#extract zip files
zip_extract()

#Homework=int(input("\nEnter homework number: "))
os.chdir(ext_dir)
data=[]
print("\nChecking submissions...")

#loop through folders
for item in os.listdir(ext_dir): 
    dir_path=os.path.join(ext_dir, item)
    
    #if item is a folder
    if os.path.isdir(dir_path):
        
        #get student id and Ex no from filename
        folder_name=item
        print("\n"+item)    #debug   
        info = folder_name.split("_")
        try:
            id = int(info[1])
            HW_no = int(info[0][2:])
        except:
            errors= errors +"--Error in Folder - "+ item + "\n"
            continue
        #print(HW_no)    #debug
        [tasks,t_req] = get_tasks(HW_no)
        tasks = int(tasks)
        t_req = list(map(int,t_req.split(";")))
        #print(t_req)    #debug
        #get author name from id no

        author = get_author(id)
        #print(author)  #debug
        #initialize variables for checks for specific student
        t_present=[0]*tasks
        t_author=[0]*tasks
        t_id=[0]*tasks
        
        #check if folder contains another folder with same name

        for folder in os.listdir(dir_path):
            print(folder +" = "+ folder_name)
            #os.path.isdir(folder) and 
            if folder == folder_name:
                print ("    Folder in Folder - "+ folder_name)

                dir_path=os.path.join(dir_path,folder_name)
                
                for folder in os.listdir(dir_path):
                    if os.path.isdir(os.path.join(dir_path,folder)) and folder == folder_name:
                        print ("    Folder in Folder in Folder - "+ folder_name)
                        errors = errors +"--Folder in Folder in Folder - "+ folder_name + "\n"
                        dir_path=os.path.join(dir_path, folder_name)
                        

        #loop through files in the folder
        for file in os.listdir(dir_path):
            print("         "+file)   #debug
            #if file is .c file
                #print (item)   #debug
            if (file.endswith(".c") or file.endswith(".C") or file.endswith(".TXT") or file.endswith(".txt")) and os.path.isfile(os.path.join(dir_path, file)):
                #get question no from file name
                t_name = os.path.basename(file)
                t_name=t_name.split(".")[0]
                print("         Checking - " + t_name)    #debug
                try:
                    t_no=int(t_name[4:])
                except:
                    errors = errors +"--Unknown file: "+dir_path+"\\"+file+"\n"
                    continue
                #print(t_no)    #debug
                
                #set presence flag
                try:
                    t_present[t_no-1]=1
                except IndexError:
                    print("Extra file in "+folder_name+"!")
                    errors = errors +"--Extra file in "+dir_path+"\\"+file+"\n"
                    continue
                    
                #open file
                check_file(file)
            else:
                errors = errors +"--Unknown file: "+dir_path+"\\"+file+"\n"
        
        #export data
        t_data=[None]*(tasks*6)
        t_grade=[0]*(tasks)     #empty for grade
        t_comments=[0]*(tasks)  #empty for comments
        t_data[::6]=t_present
        t_data[1::6]=t_req
        t_data[2::6]=t_author
        t_data[3::6]=t_id
        t_data[4::6]=t_grade    
        t_data[4::6]=t_comments 
        
        data.append([id,author,HW_no,tasks]+t_data)
        
        
print("\nChecking complete.\n\n\nExporting Report...")
#print(data)    #debug
export_data()
print(errors)
input("\n\nPress Enter to open report and exit.")
#open report
#os.startfile(program_dir+"\\AP Results.xlsm")
    


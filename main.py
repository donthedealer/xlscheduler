# Made by Liam Lawes
######################################################
#### IMPORTS
######################################################
import os
import shutil
import time as tm
import configparser
from datetime import *
import deps # deps.py is the dependency file for this program.

######################################################
#### BASE VARIABLES
######################################################
PRGMVERNO = "1.11.240625"
os.system("title XL Scheduler")
config = configparser.ConfigParser() # defines parser import

## Initial Variables and paths
forceexit = False 
filename = r"\sch-conf.ini"
user = os.getlogin()
default_cfg_directory = r"C:\Users\USER\AppData\Local\xlscheduler"
cfg_path = default_cfg_directory.replace("USER", os.getlogin())

# Define set date VAR formatted as day/month/year.
thedate = date.today().strftime('%d-%m-%Y')

######################################################
### PROGRAM FUNCTIONS
######################################################

def configini():
    
    editcfg = False
    tm.sleep(0.5)

    CFG_SECTION = "DIRECTORY"
    CFG_KEY = "scandir"
    CFG_TMPLATE = "template_file"
    default_cfg_directory = r"C:\Users\USER\AppData\Local\xlscheduler"
    cfg_path = default_cfg_directory.replace("USER", os.getlogin())

    global cfg_file
    cfg_file = cfg_path+filename

    default_cfg_contents = {
        "scandir": r"#Enter path for program to scan for folders",
        "template_file": r"#Enter file name here"
    }
    ## Actual function
    print(f"--Attempting to load config @ {cfg_path}--")
    # Attempts to find file with easy open func()
    try:
        with open(cfg_file) as fn:
            fn.close()

    except FileNotFoundError:
        print(f"ERROR: No config was found so a new config was created @ {cfg_path}.")
        print(cfg_file)
        os.system(f"mkdir {cfg_path}")
        os.system(f"echo.>{cfg_file}")
        editcfg = True
        with open(cfg_file, "w") as newcfg:
            config[CFG_SECTION] = default_cfg_contents
            config.write(newcfg)
            newcfg.close()

    if editcfg is True:
        os.system(f"start notepad.exe {cfg_file}")
        cfgpause = input("Because a new directory was created,\nPlease edit the XL scan directory, save the note and then press Enter to continue... \n")
    
    else:
        pass

    # put program_directory var on global name space load from CFG file. 
    global scan_directory
    global xl_template_file
    #global version_number

    # with open trys to find section and key
    with open(cfg_file, "r") as readfile:
        config.read(cfg_file)
        if config.has_section(CFG_SECTION) and config.has_option(CFG_SECTION, CFG_KEY) and config.has_option(CFG_SECTION, CFG_TMPLATE):
            print("--Configuration read success--")
            tm.sleep(0.5)
            scan_directory = config[CFG_SECTION][CFG_KEY]
            xl_template = config[CFG_SECTION][CFG_TMPLATE]
            xl_template_file = f"\{xl_template}"
            #version_number = config[CFG_SECTION][CFG_VER]
            #print("test" + version_number)

        # If cannot find section or key will recreate the CFG 
        else:
            readfile.close()
            with open(cfg_file, "w") as writefile:
                print("ERROR: Cannot find Directories section in config. Creating new config file...")
                config[CFG_SECTION]= default_cfg_contents
                config.write(writefile)
                scan_directory = config[CFG_SECTION][CFG_KEY]
                xl_template = config[CFG_SECTION][CFG_TMPLATE]
                
                ####print(f"You need to edit the XL Directory path and template file name in the config file at {scan_directory}")
                writefile.close()
                
    try:
        # This will move the program to operate out of the DIR loaded from config.ini
        os.chdir(scan_directory)
        print(f"--Configuration load success--\n--Program Version {PRGMVERNO}--\n")
        
    except OSError:
        print("ERROR: Unable to read directory listed in configuration, please edit configuration and/or restart program.")
        print("Opening configuration file before exiting...")
        tm.sleep(2)
        global forceexit
        forceexit = True
        os.system(f"start notepad.exe {cfg_file}")
        #os.system(f"start explorer.exe {cfg_path}")
        return forceexit

def drive_check():
    current_drive = os.getcwd()[0]
    scan_drive = scan_directory[0]
    if current_drive == scan_drive: 
        pass
    else: 
        print(f"ERROR: This program will current scan through '{scan_drive}' Drive although the program is running in '{current_drive}'.")
        print("ERROR: Please consider updating the directory with the 'chdir' command to avoid error or faulty scans.")


######################################################
### USER FUNCTIONS
######################################################

# Opens the scheduled for todays date
def get_today(): 

    # Function
    print(f"\nWould you like to open todays schedule for {thedate}?")
    while True:
        theanswer = input("Confirmation: ")
        if theanswer in ['yes', "y"]:

            # Switches to directory set to scan
            os.chdir(scan_directory)
            scanlist = []  
            folders = os.listdir()

            # Scans all folders in directory. 
            for item in folders:
                scanlist.append(item)

            # Checks current month and year
            monthstr = date.today().strftime('%B')  # Sets str var for current month.
            yearstr = date.today().strftime('%y')   # Sets str var for current year. 

            # Find folder with current month and year
            for folder in scanlist:
                if monthstr in folder:
                    if yearstr in folder: 
                        # Sets it as var
                        var = (f"\{folder}")
                        pass
                    
                else:
                    pass
                
            # Switch to this months and years folder
            try:
                folderdirectory = scan_directory + var
                os.chdir(folderdirectory)

            except UnboundLocalError:
                print("ERROR: Unable to find folder for current month in scan, try checking your scan path in the config file.\n")
                return

            # Opens todays schedule. 
            print(f"Opening your schedule for: {thedate}\n")

            try:
                os.system(f"start excel Daily_Minutes_{thedate}")
                return

            except Exception as e:
                print("ERROR: Could not open todays schedule.")
                return

        elif theanswer in ['no', 'n']:
            print()
            return
    
        else: 
            print("ERROR: Type 'no' to return to the menu.")

# Automatically copies and creates a new XL daily schedule
def createschedule():

    tyear = date.today().strftime('%Y')
    
    scanlist = []
    thescan = os.listdir()
    print("\nScanning all with this year", tyear)
    tm.sleep(1)
    # Scan initial XL Dir
    for item in thescan:
        if "template" != item and os.path.isdir(item):
            if tyear in item:
                scanlist.append(item)
    
    # check inital scan is > 0
    listvaluecheck = len(scanlist)
    if listvaluecheck == 0:
        print("ERROR: No directories found, please use 'mkdir' command and make sure the directory contains the current year.\n")
        drive_check()
        return
    
    # Scan 
    for items in scanlist:
        dirdict = dict(enumerate(scanlist))
    
    # Print each found directory
    print("Your scheduler finds the following directories:")
    print("-----------------------------------------")
    for key in dirdict:
        print(f"{key}. {dirdict[key]}")
    print("-----------------------------------------")

    while True:
        try:
            input2 = int(input("Please enter the directory number from the list: "))
            target_folder = dirdict[input2]
            print(f"Would you like to create new schedule for {target_folder}? ")
            input3 = input("Confirmation: ")
            if input3.lower() in ["y", "yes"]:
                destinpath = f"{scan_directory}\{target_folder}"
                print("Attempting to create today's schedule.")
                xl_plate = cfg_path + xl_template_file
                try:
                    shutil.copy(xl_plate, destinpath)
                    
                    
                except FileNotFoundError:
                    print(f"ERROR: The template file {xl_template_file} cannot be located in  {default_cfg_directory}\n")
                    drive_check()
                    return

                except FileExistsError:
                    print("ERROR: File already exists, no new file created.")
                    pass

                os.chdir(destinpath)

                try:
                    # Line below will remove the forward slash in set xl_template_file variable 
                    copy_xl = xl_template_file[1:]

                    # Then renames the file with the date included.
                    os.rename(copy_xl, f"Daily_Minutes_{thedate}.xlsx")

                except FileExistsError:
                    os.remove(copy_xl)
                    print(f"ERROR: There is already an XL Spreadsheet with the date: {thedate}")
                    input4 = input("Would you like to open todays schedule?: ")
                    if input4.lower() in ["y", "yes"]:
                        print(f"Opening your schedule for: {thedate}\n")
                        os.system(f"start excel Daily_Minutes_{thedate} && cd {scan_directory}")
                        drive_check()
                        break

                    elif input4.lower() in ["n", "no"]:
                        print()
                        return
                
                except PermissionError:
                    print("ERROR: Could not rename file.")
                    excelcheck = deps.process_check("EXCEL.exe")
                    if excelcheck == True:
                        print("ERROR: Microsoft Excel is already running and most likely caused an error.")
                        print("ERROR: Please close Microsoft Excel and try again.")
                    else:
                        print("PERMISSIONS-ERROR: The program could not obtain the appropriate permissions to rename the new schedule.")
                        return

                print(f"Attempting to open your schedule now...\n")
                os.system(f"start excel Daily_Minutes_{thedate}")
                drive_check()
                break

            elif input3.lower() in ["n", "no"]:
                print()
                return
            
        except ValueError:
            print("ERROR: You target a directory first, enter the number next to the directory.")
        except KeyError:
            print("ERROR: You entered a number that does not exists, please try again. Type 'no' to cancel.")


        ### WILL PROBS CUT COMMENT BELOW - TEST 30-01-2024 
        ### CUT END

    os.chdir(scan_directory)

# Simple greeting message print function
def greeting():
    time = datetime.now()
    if time.hour > 6 and time.hour < 12:
       # print(f"Good morning {os.getlogin()}")
       print(f"Good morning {user.upper()}")
    elif time.hour > 12 and time.hour < 17:
        print(f"Good afternoon {user.upper()}")
    elif time.hour > 17:
        print(f"Good evening {user.upper()}")
    print("Please type 'h' for help.\n")

# Simple function to create a new directory in the scan folder
def createnewdir():
    
    while True:
        input1 = input("\nWould you like to create a new directory?: ")
        if input1 in ["y", "yes"]:
            dirnameentry = input("Please enter the name of the new folder: ")
            # .replace() removes spaces from directory entry to avoid error
            newdir = dirnameentry.replace(" ", "")  
            print(f"Are you sure you want to create a new directory named {newdir}? ")
            confirmation = input("Confirmation: ")
            if confirmation.lower() in ["y", "yes"]:
                os.system(f"mkdir {newdir}")
                print("New Directory successfully created!")
                return

            else:
                print("ERROR: You need to enter 'yes' for confirmation.\n")
                return

        elif input1 in ["n", "no"]:
            print("Directory was NOT created!\n")
            return
                
        
        #else:
            #print("ERROR: Enter 'yes' or 'no'.")

# A function to find any schedule via user date input
def find_schedule():

    # set directory to scan dir start of every scan.
    os.chdir(scan_directory)

    year_input = input("\nPlease enter the year: ")
    if year_input in ["quit", "q", "exit"]:
        print('Exit command entered, returning to menu.\n')
        return
    month_input = input("Please enter the month: ")
    if month_input in ["quit", "q", "exit"]:
        print('Exit command entered, returning to menu.\n')
        return
    day_input = input("Please enter the day: ")
    if day_input in ["quit", "q", "exit"]:
        print('Exit command entered, returning to menu.\n')
        return

    findscan = os.listdir(scan_directory)
    userinput = f"{day_input}-{month_input}-{year_input}"

    # finddata dict will store all file names (exclusive if the directory)
    finddata = []
    pathdir = []

    for i in findscan:
        if os.path.isdir(i):
            # x all files in directories in a list 
            x = os.listdir(i)
            for f in x:
                if f.endswith(".xlsx"):
                    # Creates a list of all xlsx inclusive of dir path.     e.g. C:\scanfolder\my_file
                    pathdir.append(f"{scan_directory}\{i}\{f}")
                    # Creates a list of all Xlsx in scan directory          e.g. my_file
                    finddata.append(f)
                    
    # Creates a new list based on userinput variaable
    file_name = [i for i in finddata if userinput in i]

    #### Testing with 'pathdir' var 03-10-2023
    full_directory = [i for i in pathdir if userinput in i]  
    
    try:
        # turn new updated list of 1, into a 'result' variable.
        result = file_name[0]
        result_location = full_directory[0]
        
    except IndexError:
        print(f"ERROR: No results returned from the date {userinput}\n")
        return
    
    print("Opening Worksheet...\n")
    os.system(f"start excel {result_location}")

# Simple function to open configuration file for the program
def edit_config():

    while True:
        changedir = input("\nWould you like to edit the Configuration File?: ")
        if changedir.lower() in ["y", "yes"]:
            print("Attempting to open the configuration file...")
            os.system(f"start notepad {cfg_file}")
            tm.sleep(2)
            print("Configuration file should now be open. Returning to menu...\n")
            break

        elif changedir.lower() in ["n", "no"]:
            break

        else:
            print("ERROR: Enter 'yes' or 'no'.")

### MAIN MENU LOOP FUNCTION
def main_menu():

    mmloop = True
    while mmloop is True:
        os.chdir(scan_directory)
        mmchoice = input("TXLS COMMAND >>: ")
        if mmchoice.lower() in ["create", "c"]:
            createschedule()

        elif mmchoice.lower() in ["find",  "f"]:
            find_schedule()
        
        elif mmchoice.lower() in ["mkdir", "md"]:
            createnewdir()
            
        elif mmchoice.lower() in ["config", "configuration"]:
            edit_config()

        elif mmchoice.lower() in ["today", "tdys"]:
            get_today()

        elif mmchoice.lower() in ["clear", "cls"]:
            os.system('cls')

        elif mmchoice.lower() == "appdata": 
            os.system(f"start explorer.exe {cfg_path}")
            print(f"Opening: {cfg_path}\n")
        
        elif mmchoice.lower() in ["csd","currentscandir"]:
            print("\nThe current scan directory = ",scan_directory,"\n")

        elif mmchoice.lower() in ["quit", "q", "exit"]:
            mmloop = False

        #### Temp menu commands \/ ####
        elif mmchoice.lower() == 'pwd':
            os.system("cd")

        elif mmchoice.lower() in ["help", "h"]:
            deps.print_help()
        
        elif mmchoice.lower() in ["title"]:
            deps.print_logo()

        elif mmchoice.lower() in ["z"]:
            break
            quit 
            
        else:
            print("ERROR: Incorrect command, enter 'q' to quit or 'h' for help.")
            continue


######################################################
### PROGRAM LOOP
######################################################

deps.print_logo()   # 1. PrintS logo
configini()         # 2. Reads or Creates Configuration file in Users AppData folder
drive_check()       # 3. Checks envirnoment of terminal spawn and scan

if forceexit is False:  # 4. Loops program if force exit not triggered 
    drive_check()
    greeting()
    main_menu()

    # On Mmenu break
    print("\nHave a great day!\n")
    tm.sleep(1)

else:
    pass


######################################################
### dev notes 
######################################################
###  * Add notes here... 

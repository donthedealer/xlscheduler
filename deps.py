import subprocess

def process_check(process_name):
    sorting = [] 
    cmd = f'tasklist /FI "imagename eq {process_name}"'
    output = subprocess.check_output(cmd).decode()
    x = output.split('\n')
    y = dict(enumerate(x))
    try:
        z = y[3].split()
    except KeyError:
        print("PROCESS-ERROR: Microsoft Excel doesn't seem to be running or the function failed.")
        return
    item = z[0]
    result = item.lower().startswith(process_name.lower())
    return result


def print_logo():
    print("""                                                                         
██╗  ██╗██╗                                                              
╚██╗██╔╝██║                                                              
 ╚███╔╝ ██║                                                              
 ██╔██╗ ██║                                                              
██╔╝ ██╗███████╗                                                         
╚═╝  ╚═╝╚══════╝                                                         
                                                                         
███████╗ ██████╗██╗  ██╗███████╗██████╗ ██╗   ██╗██╗     ███████╗██████╗ 
██╔════╝██╔════╝██║  ██║██╔════╝██╔══██╗██║   ██║██║     ██╔════╝██╔══██╗
███████╗██║     ███████║█████╗  ██║  ██║██║   ██║██║     █████╗  ██████╔╝
╚════██║██║     ██╔══██║██╔══╝  ██║  ██║██║   ██║██║     ██╔══╝  ██╔══██╗
███████║╚██████╗██║  ██║███████╗██████╔╝╚██████╔╝███████╗███████╗██║  ██║
╚══════╝ ╚═════╝╚═╝  ╚═╝╚══════╝╚═════╝  ╚═════╝ ╚══════╝╚══════╝╚═╝  ╚═╝
""")

def print_help():
    print("\n--STANDARD COMMANDS--")
    print("'c/create' to Create a new schedule.")
    print("'f/find' to Find a past schedule.")   
    print("'md/mkdir' to Create a new folder within the scan directory.")
    print("'tdys/today' to Open todays schedule if already created.")
    print("'cls/clear' to Clear the contents of the window.")
    print("'q/quit' to Quit.")
    print("\n--CONFIG COMMANDS--")
    print("'csd' to Print the current scan directory." )
    print("'config/configuration' to Edit the Configuration file." )
    print("'appdata' to the application data folder.\n")

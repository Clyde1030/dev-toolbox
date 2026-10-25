@SETLOCAL ENABLEDELAYEDEXPANSION & "J:\IRPS\Team Members\Blake\EasyRunner\Python36\python.exe" -x "%~f0" %* & (IF ERRORLEVEL 1 PAUSE) & EXIT /B !ERRORLEVEL!

import os
import datetime
import subprocess
from math import floor
import git

# globals
TITLE = r"""
 _______   ________  ________       ___    ___      ________  ___  ___  ________   ________   _______   ________     
|\  ___ \ |\   __  \|\   ____\     |\  \  /  /|    |\   __  \|\  \|\  \|\   ___  \|\   ___  \|\  ___ \ |\   __  \     
\ \   __/|\ \  \|\  \ \  \___|_    \ \  \/  / /    \ \  \|\  \ \  \\\  \ \  \\ \  \ \  \\ \  \ \   __/|\ \  \|\  \    
 \ \  \_|/_\ \   __  \ \_____  \    \ \    / /      \ \   _  _\ \  \\\  \ \  \\ \  \ \  \\ \  \ \  \_|/_\ \   _  _\   
  \ \  \_|\ \ \  \ \  \|____|\  \    \/  /  /        \ \  \\  \\ \  \\\  \ \  \\ \  \ \  \\ \  \ \  \_|\ \ \  \\  \|  
   \ \_______\ \__\ \__\____\_\  \ __/  / /           \ \__\\ _\\ \_______\ \__\\ \__\ \__\\ \__\ \_______\ \__\\ _\  
    \|_______|\|__|\|__|\_________\\___/ /             \|__|\|__|\|_______|\|__| \|__|\|__| \|__|\|_______|\|__|\|__| 
"""
OPTIONS = [
    "ReinsuranceSettlements",
    "Claims",
    #"Magnastar",
    "SettlementStaging",
    "AutomaticAuditor",
    "Refresh Repos"
]
SCHEDULER_OPTIONS = [
    "repos.runRepo",
    "graph.runGraph"

]
QUARTER_OPTIONS = ("ReinsuranceSettlements", "Claims", "SettlementStaging") # options that need a quarter input
MONTH_OPTIONS = ("Magnastar", "AutomaticAuditor") # options that need a month input
DEVOPS = "pdxvmdevops01"
DEFAULT_YEAR = int(datetime.datetime.now().strftime('%Y'))
DEFAULT_MONTH = int(datetime.datetime.now().strftime('%m'))
DEFAULT_QUARTER = (floor((DEFAULT_MONTH-1)/3) - 1)%4 + 1
if DEFAULT_QUARTER == 4:
    DEFAULT_YEAR -= 1
CWD = os.getcwd()
PY_PATH = os.path.join(CWD, 'Python36', 'python.exe')

MSG = "Select what you want to run: "
YEAR_MSG = f"Year (4-digit, default {DEFAULT_YEAR}): "
MONTH_MSG = f"Month (2-digit, default {str(DEFAULT_MONTH).zfill(2)}): "
QUARTER_MSG = f"Quarter (1-digit, default {DEFAULT_QUARTER}): "

## Inputs for running the main repos ##
INPUTS = {
    "ReinsuranceSettlements": {
        "execution": """"{py_path}" \\\\{server}\\Repositories\\ReinsuranceSettlements\\main.py run {year} {quarter} --data "J:\\Acctng\\QuarterClose\\{year}\\Q{quarter}\\Assumed Settlements\\Data" --backup "\\\\{server}\\Backups\\ReinsuranceSettlements{year}Q{pq}.bak" --server {server} --root "\\\\{server}\\Repositories\\ReinsuranceSettlements" """,
        "display": """py main.py run {year} {quarter} --data "J:\\Acctng\\QuarterClose\\{year}\\Q{quarter}\\Assumed Settlements\\Data" --backup "\\\\{server}\\Backups\\ReinsuranceSettlements{year}Q{pq}.bak" """
    },
    "Claims": {
        "execution": """"{py_path}" \\\\{server}\\Repositories\\Claims\\main.py run {year} {quarter} --data "J:\\MLife\\Analysis\\Claims\\{year}\\Formatted Load Files\\{year}Q{quarter}" --backup "\\\\{server}\\Backups\\Claims{year}Q{pq}.bak" --server {server}  --root "\\\\{server}\\Repositories\\Claims" """,
        "display": """py main.py run {year} {quarter} --data "J:\\MLife\\Analysis\\Claims\\{year}\\Formatted Load Files\\{year}Q{quarter}" --backup "\\\\{server}\\Backups\\Claims{year}Q{pq}.bak" """
    },
    "SettlementStaging": {
        "execution": """"{py_path}" \\\\{server}\\Repositories\\SettlementStaging\\main.py run {year} {quarter} --backup "\\\\{server}\\Backups\\SettlementStaging{year}Q{pq}.bak"  --server {server}  --root "\\\\{server}\\Repositories\\SettlementStaging" """,
        "display": """py main.py run {year} {quarter} --backup "\\\\{server}\\Backups\\SettlementStaging{year}Q{pq}.bak" """
    },
    "AutomaticAuditor": {
        "execution": """"{py_path}" \\\\{server}\\Repositories\\AutomaticAuditor\\dashboard.py {year} {month} --report "\\\\{server}\\Repositories\\AutomaticAuditor\\runs\\dashboard.html" --root "\\\\{server}\\Repositories\\AutomaticAuditor" """,
        "display": """py dashboard.py {year} {month} --report "\\\\{server}\\Repositories\\AutomaticAuditor\\runs\\dashboard.html" """
    }
}

## Repos for running git pull ##
REPOS = [
    "ReinsuranceSettlements",
    "Claims",
    "Magnastar",
    "Exposures",
    "AutomaticAuditor",
    "SettlementStaging"
]
REPO_PATH = "\\\\{server}\\Repositories\\{repo}"

def main():
    # print title 
    print(TITLE)

    # get option choice
    for idx, opt in enumerate(OPTIONS):
        print(f"{idx+1}: {opt}")    
    
    # get the choice to run
    while True:
        try:
            choice = int(input(MSG))
        except ValueError: 
            print("Invalid input, must be number")
            continue
        
        if not (choice >= 1 and choice <= len(OPTIONS)):
            print(f"Invalid choice, must be between 1 and {len(OPTIONS)}")
            continue
        else:
            break

    # get the year and month/quarter
    option = OPTIONS[choice-1]

    if option in INPUTS.keys():
        # validation for year
        while True:
            try:
                year = int(input(YEAR_MSG) or DEFAULT_YEAR)
            except ValueError: 
                print("Invalid input, must be number")
                continue
            
            if not (year >= 2000 and year <= 2100):
                print(f"Invalid choice, must be between 2000 and 2100")
                continue
            else:
                break

        
        if option in QUARTER_OPTIONS:
            # default other option
            month = DEFAULT_MONTH
            # validation for quarter
            while True:
                try:
                    quarter = int(input(QUARTER_MSG) or DEFAULT_QUARTER)
                except ValueError: 
                    print("Invalid input, must be number")
                    continue
                
                if not (quarter >= 1 and quarter <= 4):
                    print(f"Invalid choice, must be between 1 and 4")
                    continue
                else:
                    break
        else:
            # default other option
            quarter = DEFAULT_QUARTER
            # validation for month
            while True:
                try:
                    month = int(input(MONTH_MSG) or DEFAULT_MONTH)
                except ValueError: 
                    print("Invalid input, must be number")
                    continue
                
                if not (month >= 1 and month <= 12):
                    print(f"Invalid choice, must be between 1 and 12")
                    continue
                else:
                    break

    if option in INPUTS.keys():
        executionString = INPUTS[option]['execution'].format(year=str(year), quarter=str(quarter), month=str(quarter*3).zfill(2), pq=str((quarter-2)%4 + 1), server=DEVOPS, py_path=PY_PATH)
        displayString = INPUTS[option]['display'].format(year=str(year), quarter=str(quarter), month=str(quarter*3).zfill(2), pq=str((quarter-2)%4 + 1), server=DEVOPS, py_path=PY_PATH)
        
        input(f"\nExecution string is equivalent to: {displayString}\n(Enter to execute, may take a minute to start)")
        path = os.path.join("\\\\" + DEVOPS, 'Repositories', option)        
        #os.chdir(path)
        subprocess.run(executionString, shell=False, stdout=subprocess.PIPE)
        input(f"(Completed, enter to exit)")

    elif option == 'Refresh Repos':
        # spacing for prettiness
        print('\r')

        # get option choice
        for idx, opt in enumerate(REPOS):
            print(f"{idx+1}: {opt}")    
    
        # get the choice to run
        while True:
            try:
                choice = int(input(MSG))
            except ValueError: 
                print("Invalid input, must be number")
                continue
            
            if not (choice >= 1 and choice <= len(REPOS)):
                print(f"Invalid choice, must be between 1 and {len(REPOS)}")
                continue
            else:
                break

        repo = REPOS[choice-1]
        g = git.Repo(REPO_PATH.format(server=DEVOPS, repo=repo))
        pull = g.git.pull()
        print(pull)
        input(f"(Completed, enter to exit)")
    else:
        input(f"Option has not been setup yet (enter to exit)")

    

if __name__ == "__main__":
    main()



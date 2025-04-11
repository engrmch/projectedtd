Before running the scripts, please make sure that you have installed:
*Python
*Python Libraries

to install python libraries:
Run cmd > pip install ()

pyodbc
pandas
os
fnmatch
getpass
xlwings
datetime

Step 1: CHECK IF THERE ARE INCOMPLETE DETAILS: 

Go to 'scripts' folder. Run "incompletedetails.py" script. When the script has been executed, open the "ntr.xlsx" file inside ProjectedTD folder and go to sheet "INCOMPLETE DETAILS". Endorse to Netmon those with blank TS End, negative Duration, blank Category, blank Cause. If the details are already complete or if there are no incomplete details, proceed to Step 2.

Step 2: Download the BR file from email. Put the BR file inside ProjectedTD folder.

Step 3: Run "ProjectedTDReport.bat" file. Wait until the execution of scripts is done.

Step 4: Open the TD_(currentdate).xlsx file and check if there are errors or missing data. If none, email the file.


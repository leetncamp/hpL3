# hpL3
L3 add to 9.32 report

## Quick Install

* Download this repository as a zip file or clone it using git. The main page has a link for both. 
* Next to the addL3.exe executable, place an Excel.xlsx file that contains the string "L3" in the name, e.g. "L3 IP Ranges.xlsx".  Also place an XLSX file that contains "9.32" in the name. There should only be one of each file type. 
* Do not use the older version of Excel file format .xls. Use .xlsx.  Excel can open the old type and save as the new type.
* Run addL3.exe. After it finishes, there will be a new file in the results folder. 

You should see output something like this:

<pre>
C:\Users\support2\Desktop\hpL3>addL3.exe
Loading files...

Inserting column...
Calculating networks...
Saving...C:\Users\support2\Desktop\hpL3\results\9.32 May 2018_L3.xlsx
</pre>

## Full Python Installation for Compiling hpL3.exe
### Python Install

* install git for Windows
* install python 2.7.x
* clone this repository with git clone https://github.com/leetncamp/hpL3.git hpL3
* cd into the hpL3 directory 
* run pip install -r requirements.txt
* copy your L3 Excel file into the directory
* copy your 9.32 file into the directory. 
* make sure there is only one file that contains "L3" and only one file that contains "9.32"
* python addL3.py

### Compiling a New Version

* Edit addL3.py to suit your needs
* Compile a new addL3.exe by doing this `pyinstaller.exe --onefile hpL3.py`
* This will create a "build" and a "dist" folder. Copy the new hpL3.exe file from the dist folder into the main directory.

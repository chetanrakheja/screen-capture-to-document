# Screen Capture To Document
While Testing sometime we need to take screenshots as send it to someone maybe as steps to reproduce a bug. It was very time consuming process as I had to first take the screenshot and go to word file paste that and than comeback again run application take screen shots save to doc and so on, so to decrease my efforts I created a python script that take a screen shot and save it to document


### Step 1: 
Download and install Python (v3.7+) on your system from the internet 
While installing make sure to select install pip and app python to path option

### Step 2:
once done run "pip --version" in cmd to confirm the installation
if you find a version of pip than run the following command in cmd to install the dependencies

### Step 3:
Run below command in cmd to install dependencies

**`pip install python-docx PyAutoGUI pynput docx`**

OR
```
pip install python-docx
pip install PyAutoGUI
pip install pynput
pip install docx
```
### Step 4:
save the python script to the location where you want to save the screenshots and document

### Step 5:
then in cmd go to that path and execute
"py saveImgToDoc.py"
this will run the Python script

OR just double click on saveImgToDoc.py program


**whenever you press PrtSc it will save the screenshot to a folder named shots
and to end the program and create the document press "Ctrl_l + Esc"(Left control key + Esc key)**

it will ask of the filename you want to save it as to enter the filename or just click enter, If you have entered the file name it will save as filename_12_29_52. (Filename_currentHour_currentMinute_currentSec) otherwise, by default, it will save as testingScreenShots_ currentHour_currentMinute_currentSec (Eg testingScreenShots_13_02_34.docx)

It will display a location where the screenshot is saved, where documents are saved.

Suggestions are most welcome
Email: rakhejachetan@gmail.com, chetanrakheja@gmail.com

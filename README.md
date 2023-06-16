# Stress-Score

## Introduction

The Stress Score Meter Application is a software tool designed to analyze and calculate stress scores based on data provided in an Excel sheet. 
This application aims to help organizations assess stress levels and monitor stress-related trends over time.

This tool automates the process of calculating stress scores and generating Excel sheet based
on data provided.

## Downloads

Windows setup for compiled application could be found in releases section [(here)](https://github.com/Ashutosh-Dubey2004/Stress-Score/releases).
  
Upon installation the application can be run from start menu or through desktop shortcut.

## How to run code?

Clone git repo to your local machine:  
```
git clone https://github.com/Ashutosh-Dubey2004/Stress-Score.git
```

or  

[Download this code as zip](https://github.com/Ashutosh-Dubey2004/Stress-Score/archive/refs/tags/v1.0.1.zip)  

After getting code on your local machine, follow either of the two methods.  

### Method 1:
1. Click on setup.bat to install all required python libraries.  
2. Now click on run.bat to execute the python code.  
### Method 2:
1. Open command line and type:
```cmd
pip install -r requirements.txt 
```
2. Then execute python code by typing:
```cmd
python main.py
```

### Compilation command
```
pyinstaller -F -i"images/Stress Meter.ico"  -w --noconfirm main.py --onefile --windowed
```

## Application input

* The data could be provided either in form of excel sheet (first sheet should contain the data) or csv.
* The sheet in excel file which contains data should be named 'Raw Data' (case sensitive), or the 
    first sheet in excel file should contain the data.
* First 10 columns are about student details and next 66 columns identify the survey questions or fields.
* A row must contain students' response, where the answer to survey questions could be one of the five choices:  
``` 
1
2
3
4
5
```

| Timestamp | Email Address | NAME | AGE | GENDER | MOBILE NUMBER | NAME OF INSTITUTE | PROGRAM | YEAR | ... |
| :-------- | :------------ | :--- | :-- | :----- | :------------ | :---------------- | :------ | :--- | :-- |

Required an excel or csv file where columns should be:

* TIMESTAMP
* EMAIL ADDRESS
* NAME
* AGE
* GENDER
* MOBILE NUMBER
* NAME OF INSTITUTE
* PROGRAM
* YEAR
* followed by 66 columns of questions ..

---

## Contributors:

* Ashutosh Dubey      (IMCA 2021 batch) 

Made by students at Acropolis FCA Department

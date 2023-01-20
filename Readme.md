# TalkTalk On-boarding Automation

## System Requirments
- Windows or Linux

## Dependencies
- Code
- Python
- Openpyxl 


## Downloading Dependencies
- On the top right of this page you will find a button that says **CODE**. Click on it and click **Download as a zip**
- Extract the zip file on your desktop and you will have a new folder called **TTCode-master**
- Go to https://www.python.org/downloads/ to download python
- Go to start and search `CMD`
- In command line window type `py -m pip install openpyxl` 

## Config file 
- To choose the target excel sheet you want to modify, open the config.ini file and change ``file_path`` path to the path of you excel sheet. 
- To change the name of the new editted file, open the config.ini file and change the ``updated_sheet_name``.

## How to run
- Open search and for `CMD`
- In the comamand line window type `cd code_folder_path/TTCode-master`. replace code path of the folder path where you extraced the code. 
- To run the program type `py main.py`

## Expected Results
After the code is done executing you will find a new excel sheet file in the code folder. This will be the updated sheet with the late teams coloured red and the on-time teams coloured green. 

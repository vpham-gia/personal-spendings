## Scripts to download personal spendings from Google Drive to existing Excel spreadsheet

### Setup instructions and run project
This project uses Mac OS X 10.13.4 and Python 3.5.4.
In order to run the code, two options are possible:
1. Use existing virtual environment `.venv/` folder containing installed libraries - must be in `spendings/` folder
  ```
  cd spendings/
  . activate.sh
  ```

2. Set up a similar virtual environment (could be automated using `bash init.sh`):
  1. Create a python virtual environment: `python3 -m venv .venv`
  2. Activate this newly created venv: `. activate.sh`
  3. Upgrade pip: `pip install --upgrade pip`
  4. Install requirements: `pip install -r requirements.txt`

Once the virtual environment is activated, code could be run using the following command line:  
```
python spendings/from_drive_to_xlsx/main.py
```  

### Settings and configuration
File `credentials.json` must be downloaded from https://console.developers.google.com and put in `spendings/from_drive_to_xlsx/config/` folder

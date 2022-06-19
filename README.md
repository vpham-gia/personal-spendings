## Scripts to download personal spendings from Google Drive to existing Excel spreadsheet

### Setup instructions and run project
This project uses Mac OS X 10.13.4 and Python 3.8.2.
In order to run the code, two options are possible:
1. Use existing virtual environment `.venv/` folder containing installed libraries - must be in `spendings/` folder
  ```
  cd spendings/
  . activate.sh
  ```

2. Set up a similar virtual environment 
```poetry install --no-root```

Once the virtual environment is activated, code could be run using the following command line:  
```
python spendings/from_drive_to_xlsx/main.py
```  

### Settings and configuration
File `credentials.json` must be downloaded from https://console.developers.google.com and put in `spendings/from_drive_to_xlsx/config/` folder

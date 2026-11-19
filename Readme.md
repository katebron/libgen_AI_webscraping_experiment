# RUNNING AN EXAMPLE PYTHON SCRIPT
Download the rscrape file from box – make sure to download it as a directory (aka folder)
You can run the script (get_content.py) on your local computer, but it’s probably best to make a virtual environment first – use venv, which comes with python 3.
## SETTING UP A VIRTUAL ENVIRONMENT (VENV)
Set up a venv for the rscrape folder. You can have multiple venv environments on your computer. 
- Make sure python is installed. 
 ```
 python --version 
 ```
- if on mac you may have to use
```
 python3 --version
```
 (Install python if not)
- Run in a terminal from outside the folder (you only have to do this once):
```
python -m venv rscrape
```
- Initiate the venv program (you will need to do this whenever you run the program). Run these from outside the rscrape folder 
- For mac: 
  ```
  source rscrape/bin/activate
  ```
- For windows: 
```
rscrape/Scripts/activate
```
### Install the project dependencies
 Type this command once venv is running. Cd into rscrape and type:
```
pip install -r requirements.txt
```

###	NOTE: As you move forward you may want to install your own packages – run pip package_name

- If you want to have these saved to the dependency file, run pip freeze > requirements.txt


## RUNNING THE SCRIPT
This is the easy part. I had claude write this script for me. Just run:
Pc: 
```
python get_content.py 
```
Mac:
```
 python3 get_content.py
```
You should see some output show up on screen, and then data in snapshots and a manifest csv

## STOPPING THE VIRTUAL ENVIRONMENT (VENV)
Close the terminal or run deactivate

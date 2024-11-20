### Script Description
This is a script to automatise the download of DICOM images from a visor.

### Getting Started
1. Install python in your machine. 
2. Create a virtual environment and download the requirements in `requirements.txt`.

### First Steps - Calibration of the script
1. Run the script `pyautoguiloc.py` to get the location of the mouse cursor: `python pyautoguiloc.py`
2. Modify the variables in the header **Variables to be changed** in the script `download.py` accordingly:
    a. `cerca` with the coordinates of the button to search the study
    b. `formaccession` with the coordinates of the form to write the access number
    c. `download` with the coordinates of the button to download

### Download of the images
1. Add to the folder `files` the access numbers of the studies you want to download with the name `accessnum.xlsx`. You can use the current file as a template. 
2. Run the script `download.py` with the following command: `python download.py`
3. Switch to the screen of the dicom visor

import pyautogui
import time
import json
import pandas as pd

###Variables to be changed

cerca = (856, 317)
formaccession = (1403, 394)
download = (45,34)


def search_study(access_number):
    pyautogui.moveTo(cerca)
    pyautogui.click()
    pyautogui.moveTo(formaccession)
    pyautogui.click()
    pyautogui.typewrite(access_number)
    time.sleep(1)
    pyautogui.moveTo(cerca)
    pyautogui.click()
    pyautogui.moveTo(formaccession)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'x')

def download_study():
    pyautogui.moveTo(cerca, duration=1)
    pyautogui.doubleClick()
    time.sleep(100)

def get_list(file):
    df = pd.read_excel(file)
    accession_numbers = df['accessionnumber'].astype(str).tolist()
    return accession_numbers


if __name__=="__main__":
    time.sleep(10)
    accessnum=get_list('files/accessnum.xlsx')  
    for i in accessnum:
        search_study(i)
        print(f"Searching study with accession number {i}")
        download_study()
        print(f"Downloading study with accession number {i}")



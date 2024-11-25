import pyautogui
import time
import pandas as pd
from tkinter import Tk, filedialog, messagebox, Button

# Coordinates
cerca = (914, 533)
formaccession = (653, 400)
download = (441, 630)

# Functions
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
    pyautogui.moveTo(download, duration=1)
    pyautogui.doubleClick()
    time.sleep(100)

def get_list(file):
    df = pd.read_excel(file)
    accession_numbers = df['accessionnumber'].astype(str).tolist()
    return accession_numbers

def open_file_selector():
    """Open a file dialog to select the Excel file."""
    file_path = filedialog.askopenfilename(
        title="Select the Access Number Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return None
    return file_path

def process_file():
    """Process the selected file and perform automation."""
    file_path = open_file_selector()
    if file_path:
        try:
            accessnum = get_list(file_path)
            time.sleep(5)  # Give the user time to set up the app window
            for i in accessnum:
                search_study(i)
                print(f"Searching study with accession number {i}")
                download_study()
                print(f"Downloading study with accession number {i}")
            messagebox.showinfo("Success", "All studies processed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Main Script with GUI
if __name__ == "__main__":
    root = Tk()
    root.title("Download Dicom Images")
    root.geometry("300x150")
    root.resizable(False, False)

    # Add a button to trigger file selection
    button = Button(
        root,
        text="Select File and Start",
        command=process_file,
        width=20,
        height=2
    )
    button.pack(pady=40)

    # Start the Tkinter event loop
    root.mainloop()

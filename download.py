import pyautogui
import time
import pandas as pd
from tkinter import *
from tkinter import messagebox, filedialog
from tkinter import ttk

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
    time.sleep(30)

def get_list(file):
    df = pd.read_excel(file)
    accession_numbers = df['accessionnumber'].astype(str).tolist()
    return accession_numbers

def open_file_selector():
    """Open a file dialog to select the Excel file."""
    file_path = filedialog.askopenfilename(
        title="Selecciona l'excel amb els accession numbers",
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
            root.destroy()
            messagebox.showinfo("Success", "Clica OK i vés a StarViewer a la pantalla del PACS")
            accessnum = get_list(file_path)
            time.sleep(7)  # Give the user time to set up the app window
            for i in accessnum:
                search_study(i)
                print(f"Searching study with accession number {i}")
                download_study()
                print(f"Downloading study with accession number {i}")
            messagebox.showinfo("Success", "S'han baixat totes les imatges")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

# Main Script with GUI
if __name__ == "__main__":
    root = Tk()
    root.title("Download DICOM Images")
    root.geometry("550x150")  

    mainframe = ttk.Frame(root, padding="3 3 12 12")
    mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    ttk.Label(mainframe, text="Selecciona el excel que conté els access numbers.").grid(column=3, row=1, sticky=(S))
    ttk.Label(mainframe, text="Ha de seguir el mateix format que l'excel de la carpeta de la app.").grid(column=3, row=2, sticky=(S))
    ttk.Button(mainframe, text="Select File", command=process_file).grid(column=3, row=3, sticky=(S))
    ttk.Label(mainframe, text="Created by Xabier Michelena").grid(column=3, row=4, sticky=(S))

    root.mainloop()

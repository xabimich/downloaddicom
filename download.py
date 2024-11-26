import pyautogui
import time
import pandas as pd
from tkinter import *
from tkinter import messagebox, filedialog
from tkinter import ttk
import threading

# PyAutoGUI failsafe - move mouse to corner to stop the script
pyautogui.FAILSAFE = True

# Coordinates
cerca = (855, 499)
formaccession = (640, 363)
download = (242, 590)

# Global variable to control the process
should_stop = False

def search_study(access_number):
    pyautogui.moveTo(formaccession, duration=0.5)
    pyautogui.click()
    pyautogui.typewrite(access_number, interval=0.1)
    time.sleep(1)
    pyautogui.moveTo(cerca, duration=0.5)
    pyautogui.click()
    pyautogui.moveTo(formaccession, duration=0.5)
    pyautogui.click()
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'x')

def download_study():
    pyautogui.moveTo(download, duration=3)
    pyautogui.doubleClick()
    time.sleep(30)

def get_list(file):
    df = pd.read_excel(file)
    accession_numbers = df['accessionnumber'].astype(str).tolist()
    return accession_numbers

def open_file_selector():
    file_path = filedialog.askopenfilename(
        title="Selecciona l'excel amb els accession numbers",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return None
    return file_path

def stop_process():
    global should_stop
    should_stop = True
    messagebox.showinfo("Stop", "El procés ha parat")

def create_control_window():
    """Update the GUI with control elements."""
    for widget in mainframe.winfo_children():
        widget.destroy()  # Clear all widgets in mainframe

    ttk.Label(mainframe, text="El procés de baixar imatges està funcionant").grid(column=3, row=1, sticky=(S))
    ttk.Button(mainframe, text="Parar procés", command=stop_process).grid(column=3, row=2, sticky=(S))
    ttk.Label(mainframe, text="Created by Xabier Michelena").grid(column=3, row=3, sticky=(S))

def process_file_in_thread(file_path):
    """Process the file in a separate thread."""
    global should_stop
    should_stop = False  # Reset the stop flag
    try:
        messagebox.showinfo("Success", "Clica OK i vés a StarViewer a la pantalla del PACS")
        accessnum = get_list(file_path)
        time.sleep(7)  # Time for user to set up the app window
        create_control_window()  # Update the control window
        for i in accessnum:
            if should_stop:
                print("Procés finalitzat")
                break
            search_study(i)
            print(f"Searching study with accession number {i}")
            download_study()
            print(f"Downloading study with accession number {i}")
        if not should_stop:
            messagebox.showinfo("Success", "S'han baixat totes les imatges")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def process_file():
    """Start file processing."""
    file_path = open_file_selector()
    if file_path:
        threading.Thread(target=process_file_in_thread, args=(file_path,), daemon=True).start()

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

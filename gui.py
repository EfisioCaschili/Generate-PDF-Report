import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
import subprocess
import sys
import os
try:
    from dotenv import dotenv_values, load_dotenv
except:
    from dotenv import main

path=f"{str(os.getcwd())}\\"  
#path= os.path.dirname(__file__)
os.chdir(os.path.dirname(os.path.abspath(__file__)))
try:
    env=dotenv_values(path+'env.env')
    dest_file=env.get('dest_file')
except: 
    env=main.load_dotenv(path+'env.env')
    dest_file=os.getenv('dest_file')

def change_date():
    today = str(entry_start_date.get_date())
    tomorrow = str(entry_end_date.get_date())
    
    f=open(dest_file+"isloaded.txt", 'w')
    f.close()
    result = subprocess.run([sys.executable, 'main.py', '--today', today, '--tomorrow', tomorrow], capture_output=True)
    
    # Mostra l'output nello stesso Tkinter Text Widget
    output_text.delete(1.0, tk.END)  # Pulisci il widget di testo
    output_text.insert(tk.END, f"Output:\n{result.stdout.decode('utf-8')}\n Errors:\n{result.stderr.decode('utf-8')}")
    if "holidays" in str(result.stdout):
         return
    """ with open(dest_file+"isloaded.txt", 'w') as file:#create again the file isloaded.txt
            file.write(today)"""

# Creazione della finestra principale
root = tk.Tk()
root.title("GUI")
root.iconbitmap()
root.iconbitmap("//192.168.1.125/gbts/Create_PDF_Report/images/ajt_official.ico")

# Campi di inserimento per le date
ttk.Label(root, text="Today:").grid(row=0, column=0, padx=5, pady=5) 
entry_start_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_start_date.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(root, text="Tomorrow:").grid(row=1, column=0, padx=5, pady=5) 
entry_end_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
entry_end_date.grid(row=1, column=1, padx=5, pady=5)

# Pulsante per eseguire lo script
run_button = ttk.Button(root, text="Generate Report ", command=change_date) 
run_button.grid(row=2, column=0, columnspan=2, pady=10)


# Widget di testo per l'output
output_text = tk.Text(root, height=10, width=50) 
output_text.grid(row=3, column=0, columnspan=2, pady=10)

root.mainloop()



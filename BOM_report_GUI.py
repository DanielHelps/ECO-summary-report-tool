import tkinter as tk
from tkinter import ttk
import os
import BOM_report
import tkinter.messagebox
import pickle
from requests.auth import HTTPBasicAuth
import xlsxwriter
import os
from tkinter import StringVar
from functools import partial   
import sys

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def save_nd_quit():
    config = {
        'username': username.get(),
        'password': password.get(),
    }

    with open("saved_settings.dat", "wb") as pickle_file:
        pickle.dump(config, pickle_file, pickle.HIGHEST_PROTOCOL)
    root.quit()
    

def load():
    with open("saved_settings.dat", "rb") as pickle_file:
        config = pickle.load(pickle_file)

    username.set(config.get('username'))
    password.set(config.get('password'))

root = tk.Tk()
root.geometry('420x230')
root.resizable(False, False)
root.title('ECO summary report tool V1.2')
title_font = ("Helvetica", 16)
regular_font = ("Helvetica", 12)
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.iconbitmap(resource_path("icon.ico"))

def BOM_report_creation(username, password, ECO_num):

    # tkinter.messagebox.showerror("ERROR!", f'Username {username.get()}, Password {password.get()}, ECO # {ECO_num.get()}')
    
    # BOM_report.main(username.get(),password.get(),ECO_num.get())

    # username = 'daniel.marom@kornit.com'
    # password = 'Kornit@2023'
    # ECO_num = 'ECO-10029-23'

   

    if ECO_num.get() != "":
        try:
            response_code = BOM_report.main(username.get(),password.get(),ECO_num.get(), pb, value_label, root)
        except PermissionError:
            tkinter.messagebox.showerror("Permission error", "Permission error, please exit the open ECO report and try again")
        except IndexError:
            tkinter.messagebox.showerror("ECO error", "Couldn't find ECO, please check ECO #")
        except ZeroDivisionError:
            tkinter.messagebox.showerror("ECO error", "0 items in ECO... Cannot create ECO report")
        except Exception as e:
            tkinter.messagebox.showerror("Unkown error", e)
        else:
            if response_code==401:
                tkinter.messagebox.showerror("Authorization error", "Authorization error, please check your username and password!")
            else:
                tkinter.messagebox.showinfo("Report complete", f"""Report complete! 
you can find the it in the current directory at 
   {response_code}""")
    else:
        tkinter.messagebox.showerror("ECO error", "Can't have empty ECO #")
            


ttk.Label(
    root,
    text='Oracle username',
    font=title_font).grid(row=0, column=0, sticky=tk.W, padx=10)

username = StringVar()
usernameEntry = ttk.Entry(root, textvariable=username).grid(row=0, column=1, sticky=tk.W, ipadx=40, padx=(0,15))


ttk.Label(
    root,
    text='Oracle password',
    font=title_font).grid(row=1, column=0, sticky=tk.W, padx=10)

password = StringVar()
passwordEntry = ttk.Entry(root, textvariable=password, show='*').grid(row=1, column=1, sticky=tk.W, ipadx=40, padx=(0,15))  


if os.path.isfile(f'saved_settings.dat'):
    load()

ttk.Label(
    root,
    text='ECO #',
    font=title_font).grid(row=2, column=0, sticky=tk.W, padx=12)

ECO_num = StringVar()
ecoEntry = ttk.Entry(root, textvariable=ECO_num).grid(row=2, column=1, sticky=tk.W, ipadx=40, padx=(0,15))  


pb = ttk.Progressbar(
    root,
    orient='horizontal',
    mode='determinate',
    length=320
)
pb.grid(row=3, column=0, columnspan=2, padx=10, pady=30, sticky=tk.W)
value_label = tk.Label(root, text="0%")
value_label.grid(row=3, column=1, padx=20, columnspan=2, sticky=tk.E)

BOM_report_func = partial(BOM_report_creation, username, password, ECO_num)

create_Report_But = ttk.Button(root, text="Create report", command=BOM_report_func).grid(row=4, column=0, ipadx=20, ipady=10)  

exit_but = ttk.Button(root, text="Exit", command=save_nd_quit)
exit_but.grid(row=4, column=1, ipadx=20, ipady=10)

def show(event=None): # handler
    BOM_report_creation(username,password,ECO_num)


root.bind('<Return>', show)

root.mainloop()
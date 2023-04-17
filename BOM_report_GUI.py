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
from check_conflicts import create_conflict_report
import urllib3
import keyring

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def save_nd_quit():
    # config = {
    #     'username': username.get(),
    #     'password': password.get(),
    # }

    # with open("saved_settings.dat", "wb") as pickle_file:
    #     pickle.dump(config, pickle_file, pickle.HIGHEST_PROTOCOL)

    config = {
        'username': username.get(),
    }
    if not os.path.exists(os.path.expanduser('~\\AppData\\Local\\BOM_report')):
        os.makedirs(os.path.expanduser('~\\AppData\\Local\\BOM_report'))

    with open(os.path.expanduser('~\\AppData\\Local\\BOM_report\\saved_settings.dat'), "wb") as pickle_file:
        pickle.dump(config, pickle_file, pickle.HIGHEST_PROTOCOL)

    keyring.set_password("BOM_report", username.get(), password.get())
    root.quit()
    

def load():
    if os.path.exists(os.path.expanduser('~\\AppData\\Local\\BOM_report\\saved_settings.dat')):
        with open(os.path.expanduser('~\\AppData\\Local\\BOM_report\\saved_settings.dat'), "rb") as pickle_file:
            config = pickle.load(pickle_file)

        username.set(config.get('username'))
        # password.set(config.get('password'))

        password.set(keyring.get_password("BOM_report",config.get('username')))

    


urllib3.disable_warnings()
root = tk.Tk()
root.geometry('420x280')
root.resizable(False, False)
root.title('ECO summary report tool V2.1')
title_font = ("Helvetica", 16)
regular_font = ("Helvetica", 12)
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.iconbitmap(resource_path("icon.ico"))

def Report_creation(username, password, ECO_num, report_code):


    if ECO_num.get() != "":
        try:
            if report_code == 1:
                response_code = BOM_report.main(username.get(),password.get(),ECO_num.get(), pb, value_label, root)
            else:
                response_code = create_conflict_report(username.get(),password.get(),ECO_num.get(), pb, value_label, root)
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

BOM_report_func = partial(Report_creation, username, password, ECO_num, 1)
Conflict_report_func = partial(Report_creation, username, password, ECO_num, 2)

create_Report_But = ttk.Button(root, text="""    Create ECO
summary report""", command=BOM_report_func).grid(row=4, column=0, ipadx=20, ipady=10)  

create_Conflict_But = ttk.Button(root, text="""    Create ECO
conflict report""", command=Conflict_report_func).grid(row=5, column=0, ipadx=20, pady=10)  

# create_Report_But = ttk.Button(root, text="Create report", command=BOM_report_func).grid(row=4, column=0, ipadx=20, ipady=10)  

# create_Report_But = ttk.Button(root, text="Create report", command=BOM_report_func).grid(row=4, column=0, ipadx=20, ipady=10)  

exit_but = ttk.Button(root, text="Exit", command=save_nd_quit)
exit_but.grid(row=5, column=1, ipadx=20, ipady=10)

def show(event=None): # handler
    Report_creation(username,password,ECO_num)


root.bind('<Return>', show)

root.mainloop()
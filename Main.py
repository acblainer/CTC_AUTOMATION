from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import tkinter.font as font
from tkinter import messagebox
import threading
from threading import Thread #to extend Thread class
from CTC_Automation_Package import Selling_Curve_N_Consolidation
from CTC_Automation_Package import RecShipFinal_Adam
import pkg_resources
import subprocess
import sys
import os
import sqlite3

#install the prerequire
def gui_prep():
    def install(package):
        if (package not in {pkg.key for pkg in pkg_resources.working_set}):
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
    for pkg in ['pandas', 'numpy', 'pillow']:
        install(pkg)

gui_prep()
from PIL import ImageTk, Image


#start the root window
root = Tk()
#give a title of the root window
root.title("CTC Automation Tools")
#provide a default size of the window
root.geometry("600x500")
#https://stackoverflow.com/questions/45847313/what-does-weight-do-in-tkinter
root.grid_columnconfigure(0,weight = 1)
root.grid_columnconfigure(1,weight = 1)
root.grid_columnconfigure(2,weight = 1)
root.grid_columnconfigure(3,weight = 1)

##create the entry widget, this is mainly to use as an indicator
e = Entry(root, width = 100, borderwidth=3, justify = CENTER)
e.config(state = 'normal')
e.insert(0, "This is to show you which tool you have chosen!")
e.config(state = 'disabled')
e.grid(row = 0, column = 1, columnspan = 3, padx = 5, pady = 20, sticky = EW)

#add logo to the app
logo = ImageTk.PhotoImage(Image.open('./Image/logo.png').resize((170,30)))
logo_label = Label(image = logo)
logo_label.grid(row = 0, column = 0)

#then I will create a progression bar for any tool being used
progress_bar = ttk.Progressbar(root, orient = HORIZONTAL, length = 150, mode = 'indeterminate')

#create a frame to hold the parameter section
#however, this window will only fire when the user has not yet correctly enter the username and password
#I will create a sqlite database to hold the credentials for Orachle DB connection
#this sqlite database will site in my local file
#when you connect to an SQLite database that does not exist, SQLite automatically creates the new database for you
def create_sql_lite_conn(db_file, **kargs):
    '''create a database connection to a SQLite database'''
    conn = None
    cur = None
    try:
        conn = sqlite3.connect(db_file)
        cur = conn.cursor()
        #if there is no table exists to holde the credentials, I will create a table
        #my os.getlogin() is Yongpeng.Fu for example. One computer one username and password to use
        cur.execute('''create table if not exists dblite 
                    (HOSTNAME TEXT, DB TEXT, DIALECT TEXT, SQL_DRIVER TEXT, USERNAME TEXT, PASSWORD TEXT, HOST TEXT, PORT INTEGER, SERVICE TEXT)''')
        conn.commit()
        if 'USERNAME' in kargs:
            cur.execute('''INSERT INTO dblite VALUES (?,?,?,?,?,?,?,?,?)''',
            (os.getlogin(), kargs["DB"], 'oracle', 'cx_oracle', kargs["USERNAME"], kargs["PASSWORD"],"p9cpwpjdadb01",25959, "FR01PR"))
            conn.commit()
        else:
            cur.execute('''SELECT * FROM dblite WHERE HOSTNAME = ? and DB = ?''', (os.getlogin(),kargs["DB"]))
            return cur.fetchone()
    except sqlite3.Error as e:
        messagebox.showerror(title = "Something Wrong", message = e)
    finally:
        if conn:
            cur.close()
            conn.close()


#Extend Thread class so that I can decide if I want to write the credential to the sqlite database or not
#https://stackoverflow.com/questions/2829329/catch-a-threads-exception-in-the-caller-thread
#https://stackoverflow.com/questions/34742026/python-pass-argument-to-thread-class
class Capture_Child_Thread_Error(Thread):
    def __init__(self, func, db_indicator, *args):
        #daemon= True tells this thread to end as well when the caller thread is killed
        Thread.__init__(self, daemon=True)
        self.args = args
        self.func =func
        self.db_indicator = db_indicator
    #override the run function
    def run(self):
        try:
            self.func(self.args[0],self.args[1])
        except Exception as err:
            messagebox.showerror(title = "Something Wrong", message = err)
        else:
            messagebox.showinfo(title = "Result", message = "Job Done!")
            if self.db_indicator:
                #if nothing is wrong, write this credential to the database
                create_sql_lite_conn('./access.db', DB = "selling", USERNAME = user_name, PASSWORD = password)
        finally:
            #when the code comes here, clear out the progress bar
            progress_bar.stop()
            #after the job is done remove the progression bar
            progress_bar.grid_forget()
            file_location_input.delete(0,END)
            e.config(state = 'normal')
            e.delete(0, END)
            e.insert(0, "This is to show you which tool you have chosen!")
            e.config(state = 'disabled')

#define what will happen when you click OK button
def ok_btn_func():
    global user_name
    global password
    user_name = user_entry.get()
    password = pass_entry.get()
    top.destroy()
    #put the progression bar in there
    progress_bar.grid(row = 5, column = 1, columnspan = 2, padx = 10, pady = 20)
    progress_bar.start(18)
    selling_para_dict = {'DIALECT':'oracle', 'SQL_DRIVER':'cx_oracle', 'USERNAME':user_name, 
                                    'PASSWORD':password, 'HOST':'p9cpwpjdadb01', 'PORT':25959,'SERVICE':'FR01PR'}
    #since you already come to the page to ask you enter, I want to store this credential into the DB, so True
    thread_selling_curve = Capture_Child_Thread_Error(Selling_Curve, True, file_location_input.get(), selling_para_dict)
    thread_selling_curve.start()
    # threading.Thread(target = Selling_Curve, args = (file_location_input.get(), selling_para_dict), daemon= True).start()

#Define a function to accecpt username and password if necessary
def para_window(root):
    #open a new window
    global top
    top = Toplevel(root)
    top.title("Credential Window for Database")
    top.geometry("350x250")
    #username label
    global user_entry
    user_label = Label(top, text = "USERNAME:", pady = 50, padx = 5)
    user_label.grid(row = 0, column = 0)
    user_entry = Entry(top, bd = 3, width = 35)
    user_entry.grid(row = 0, column = 1)
    #password label
    global pass_entry
    pass_label = Label(top, text = "PASSWORD:", pady = 20, padx = 5)
    pass_label.grid(row = 1, column = 0)
    pass_entry = Entry(top, bd = 3,width = 35, show = "*")
    pass_entry.grid(row = 1, column = 1)
    #define a OK button
    ok_btn = Button(top, text = "OK", command = ok_btn_func, width = 5)
    ok_btn.grid(row = 2, column = 1)
    #disable the underlying window when a second window pops up
    # top.wait_visibility()
    # top.grab_set_global()

#define a function to actually run the work for RecShip
def ok_btn_style_func():
    #put the progression bar in there
    progress_bar.grid(row = 5, column = 1, columnspan = 2, padx = 10, pady = 20)
    progress_bar.start(18)
    adam_user_name = None
    adam_password = None
    recship_para_dict = {'DIALECT':'oracle', 'SQL_DRIVER':'cx_oracle', 'USERNAME':adam_user_name, 
                        'PASSWORD':adam_password, 'HOST':'p9cpwpjdadb01', 'PORT':25959,'SERVICE':'FR01PR'}
    top_style.destroy()

#define a function to accept a list of styles
def style_window(root):
    #open a new window
    global top_style
    top_style = Toplevel(root)
    top_style.title("A list of styles")
    top_style.geometry("600x300")
    #username label
    global list_entry
    list_label = Label(top_style, text = "A list of styles:")
    #default list of styles
    default_styles = "A list of styles with singe quotation mark, and separated by , as follows:\n 'S_6AREDHASHBSPTS2', 'S_75729', 'S_79645', 'S_79646'"
    list_text = Text(top_style, height = 10, width = 52)
    #define a OK button
    ok_btn = Button(top_style, text = "OK", command = ok_btn_style_func, width = 5)
    #put them on the window
    list_label.pack()
    list_text.pack(fill = BOTH)
    ok_btn.pack()
    #provide a default look of list
    list_text.insert(END, default_styles)


#funtion to do the selling curve work, the function to be used in another thread
def Selling_Curve(file_path, kargs):
    Selling_Curve_N_Consolidation.selling_curve_prep()
    #I will pass a dictionary to the Selling_Curve but then unpack it in the read_query_output
    Selling_Curve_N_Consolidation.read_query_output(file_path, **kargs)

#funtion to do the selling curve work, the function to be used in another thread
def consolidation_func(file_path):
    try:
        Selling_Curve_N_Consolidation.consolidation_prep()
        Selling_Curve_N_Consolidation.consolidation_output(file_path)
    except Exception as err:
        messagebox.showerror(title = "Something Wrong", message = err)
    else:
        messagebox.showinfo(title = "Result", message = "Job Done!")
    finally:
        progress_bar.stop()
        #after the job is done remove the progression bar
        progress_bar.grid_forget()
        file_location_input.delete(0,END)
        e.config(state = 'normal')
        e.delete(0, END)
        e.insert(0, "This is to show you which tool you have chosen!")
        e.config(state = 'disabled')


#define a function to update the text inside entry widget
def button_click(text_button):
    e.config(state = 'normal')
    e.delete(0, END)
    e.insert(0, text_button)
    e.config(state = 'disabled')
    if text_button.split()[0].lower() == "selling":
        file_path = filedialog.askopenfilename()
        file_location_input.insert(0,file_path)
        ##if this funciton returns something then we know the connection is success, 
        # #we will use the info to connect oracle, and not fire up the parameter window
        credential_sqlite = create_sql_lite_conn('./access.db',DB = "selling")
        if credential_sqlite:
            #put the progression bar in there
            progress_bar.grid(row = 5, column = 1, columnspan = 2, padx = 10, pady = 20)
            progress_bar.start(18)
            #start the real job in another thread
            selling_para_dict = {'DIALECT':credential_sqlite[2], 'SQL_DRIVER':credential_sqlite[3], 'USERNAME':credential_sqlite[4], 
                                    'PASSWORD':credential_sqlite[5], 'HOST':credential_sqlite[6], 'PORT':credential_sqlite[7],'SERVICE':credential_sqlite[8]}
            #because the credential is already found in the database, there is no need to write the info into database, so False
            thread_selling_curve = Capture_Child_Thread_Error(Selling_Curve, False, file_location_input.get(), selling_para_dict)
            thread_selling_curve.start()
            # threading.Thread(target = Selling_Curve, args = (file_location_input.get(), selling_para_dict), daemon= True).start()
        else:
            #fire up the parameter window
            para_window(root)
                

    if text_button.split()[0].lower() == "consolidation":
        file_path = filedialog.askopenfilename()
        file_location_input.insert(0,file_path)
        #put the progression bar in there
        progress_bar.grid(row = 5, column = 1, columnspan = 2, padx = 10, pady = 20)
        progress_bar.start(18)
        #start the real job in another thread, daemon= True tells the thread to shut down as well when the main thread is off
        threading.Thread(target = consolidation_func, args = (file_location_input.get(),), daemon= True).start()


    if text_button.split()[0].lower() == "recship":
        style_window(root)

    

#create 2 buttons one for Selling Curve and another one for Consolidaiton
button_selling = Button(root, text = "Selling Curve Tool", borderwidth = 3, padx = 70, pady = 40, command = lambda:button_click('Selling Curve Tool Selected'))
button_selling.grid(row = 2, column = 0, columnspan = 2, sticky=NSEW)
button_selling['font'] = font.Font(size = 15)
Consolidaiton = Button(root, text = "Consolidation Tool", borderwidth = 3, padx = 70, pady = 40,command = lambda:button_click('Consolidation Tool Selected'))
Consolidaiton.grid(row = 2, column = 2,columnspan = 2, sticky=NSEW)
Consolidaiton['font'] = font.Font(size = 15)
#creat another button to run Adam's code
button_recship = Button(root, text = "RecShip Tool", borderwidth = 3, padx = 70, pady = 40,command = lambda:button_click('RecShip Tool Selected'))
button_recship.grid(row = 3, column = 0, columnspan = 4, sticky=NSEW)
button_recship['font'] = font.Font(size = 15)
#parameters for each tool, you need the file path at least
file_location = Label(root, text = "File Location:")
file_location.grid(row = 4, column = 0)
file_location_input = Entry(root, width = 100, borderwidth=3, justify = CENTER)
file_location_input.grid(row = 4, column = 1, columnspan = 3, pady = 20, sticky = EW)

#creat the event loop for the GUI
root.mainloop()
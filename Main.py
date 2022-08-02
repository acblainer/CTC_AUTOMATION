from tkinter import *
from tkinter import filedialog
from tkinter import ttk
import tkinter.font as font
from tkinter import messagebox
import threading
from CTC_Automation_Package import Selling_Curve_N_Consolidation
import pkg_resources
import subprocess
import sys

#install the prerequire
def gui_prep():
    def install(package):
        if (package not in {pkg.key for pkg in pkg_resources.working_set}):
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
    for pkg in ['pandas', 'numpy', 'pillow']:
        install(pkg)

gui_prep()
from PIL import ImageTk, Image


root = Tk()
#give a title of the root window
root.title("CTC Automation Tools")
#provide a default size of the window
root.geometry("600x400")
root.grid_columnconfigure(0,weight = 1)
root.grid_columnconfigure(1,weight = 1)
root.grid_columnconfigure(2,weight = 1)
root.grid_columnconfigure(3,weight = 1)

##create the entry widget, this is mainly to use as an indicator
e = Entry(root, width = 100, borderwidth=3, justify = CENTER)
e.config(state = 'normal')
e.insert(0, "This is to show you which tool you have chosen!")
e.config(state = 'disabled')
e.grid(row = 0, column = 1, columnspan = 3, padx = 5, pady = 20, sticky = E)

#add logo to the app
logo = ImageTk.PhotoImage(Image.open('./Image/logo.png').resize((170,30)))
logo_label = Label(image = logo)
logo_label.grid(row = 0, column = 0)

#then I will create a progression bar for any tool being used
progress_bar = ttk.Progressbar(root, orient = HORIZONTAL, length = 150, mode = 'indeterminate')

#funtion to do the selling curve work, the function to be used in another thread
def Selling_Curve(file_path):
    try:
        Selling_Curve_N_Consolidation.selling_curve_prep()
        Selling_Curve_N_Consolidation.read_query_output(file_path)
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
        #put the progression bar in there
        progress_bar.grid(row = 4, column = 1, columnspan = 2, padx = 10, pady = 20)
        progress_bar.start(18)
        #start the real job in another thread
        threading.Thread(target = Selling_Curve, args = (file_location_input.get(),)).start()

    if text_button.split()[0].lower() == "consolidation":
        file_path = filedialog.askopenfilename()
        file_location_input.insert(0,file_path)
        #put the progression bar in there
        progress_bar.grid(row = 4, column = 1, columnspan = 2, padx = 10, pady = 20)
        progress_bar.start(18)
        #start the real job in another thread
        threading.Thread(target = consolidation_func, args = (file_location_input.get(),)).start()

#create 2 buttons one for Selling Curve and another one for Consolidaiton
button_selling = Button(root, text = "Selling Curve Tool", borderwidth = 3, padx = 70, pady = 80, command = lambda:button_click('Selling Curve Tool Selected'))
button_selling.grid(row = 2, column = 0, columnspan = 2, sticky=NSEW)
button_selling['font'] = font.Font(size = 15)
Consolidaiton = Button(root, text = "Consolidation Tool", borderwidth = 3, padx = 70, pady = 80,command = lambda:button_click('Consolidation Tool Selected'))
Consolidaiton.grid(row = 2, column = 2,columnspan = 2, sticky=NSEW)
Consolidaiton['font'] = font.Font(size = 15)
#parameters for each tool, you need the file path at least
file_location = Label(root, text = "File Location:")
file_location.grid(row = 3, column = 0)
file_location_input = Entry(root, width = 100, borderwidth=3, justify = CENTER)
file_location_input.grid(row = 3, column = 1, columnspan = 3, pady = 20, sticky = E)
#get a trade mark 
# trade = Label(root, text = "Mark's", bd = 1, relief = SUNKEN, anchor = E)
# trade.grid(row = 4, column = 0, columnspan=4, sticky = EW)
#create a frame to hold the parameter section
def create_new_window():
    win = Toplevel(root)

para_frame = LabelFrame(root, text = 'parameter')
para_frame.grid(row = 5, column = 0)
place_holder = Label(para_frame, text = "place holder")
place_holder.pack()



#creat the event loop for the GUI
root.mainloop()
from tkinter import *

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
e.grid(row = 0, column = 0, columnspan = 4, padx = 5, pady = 20)

#define a function to update the text inside entry widget
def button_click(text_button):
    e.config(state = 'normal')
    e.delete(0, END)
    e.insert(0, text_button)
    e.config(state = 'disabled')
    if text_button.split()[0].lower() == "selling":
        pass
    if text_button.split()[0].lower() == "consolidation":
        pass

#create 2 buttons one for Selling Curve and another one for Consolidaiton
button_selling = Button(root, text = "Selling Curve Tool", borderwidth = 3, padx = 70, pady = 50, command = lambda:button_click('Selling Curve Tool Selected'))
button_selling.grid(row = 2, column = 0, columnspan = 2, sticky=NSEW)
Consolidaiton = Button(root, text = "Consolidation Tool", borderwidth = 3, padx = 70, pady = 50,command = lambda:button_click('Consolidation Tool Selected'))
Consolidaiton.grid(row = 2, column = 2,columnspan = 2, sticky=NSEW)
#parameters for each tool
#you need the file path

root.mainloop()
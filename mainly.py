import os
from subprocess import call

import sys

try:
    from Tkinter import *
except ImportError:
    from tkinter import *

try:
    import ttk
    py3 = False
except ImportError:
    import tkinter.ttk as ttk
    py3 = True
def click_checkinn():
    call(["python", "booking.py"])
def click_list():
    call(["python", "list.py"])
def click_checkout():
    call(["python", "checkout.py"])
def click_getinfo():
    call(["python","getinfo.py"])


class HOMESTAY_MANAGEMENT:
    def __init__(self):
        root = Tk()
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#062830'  
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#ffffff' # X11 color: 'white'
        _ana1color = '#ffffff' # X11 color: 'white'
        _ana2color = '#ffffff' # X11 color: 'white'
        font14 = "-family {Segoe UI} -size 15 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font16 = "-family {Swis721 BlkCn BT} -size 40 -weight bold "  \
            "-slant roman -underline 0 -overstrike 0"
        font9 = "-family {Segoe UI} -size 9 -weight normal -slant "  \
            "roman -underline 0 -overstrike 0"

        root.geometry("963x749+540+110")
        root.title("LA CASA HOMESTAY BOOKING SYSTEM")
        root.configure(background="#062830")
        root.configure(highlightbackground="#d9d9d9")
        root.configure(highlightcolor="black")



        self.menubar = Menu(root,font=font9,bg=_bgcolor,fg=_fgcolor)
        root.configure(menu = self.menubar)



        self.Frame1 = Frame(root)
        self.Frame1.place(relx=0.02, rely=0.03, relheight=0.94, relwidth=0.96)
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(borderwidth="2")
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(background="beige")
        self.Frame1.configure(highlightbackground="#d9d9d9")
        self.Frame1.configure(highlightcolor="black")
        self.Frame1.configure(width=925)

        self.Message6 = Message(self.Frame1)
        self.Message6.place(relx=0.09, rely=0.01, relheight=0.15, relwidth=0.90)
        self.Message6.configure(background="beige")
        self.Message6.configure(font=font16)
        self.Message6.configure(foreground="#000000")
        self.Message6.configure(highlightbackground="#d9d9d9")
        self.Message6.configure(highlightcolor="black")
        self.Message6.configure(text='''LA CASA HOMESTAY''')
        self.Message6.configure(width=891)

        self.Button2 = Button(self.Frame1)
        self.Button2.place(relx=0.15, rely=0.20, height=85, width=300)
        self.Button2.configure(activebackground="beige")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#062830")
        self.Button2.configure(disabledforeground="#bfbfbf")
        self.Button2.configure(font=font14)
        self.Button2.configure(foreground="white")
        self.Button2.configure(highlightbackground="#d9d9d9")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''BOOKING''')
        self.Button2.configure(width=566)
        self.Button2.configure(command=click_checkinn)

        self.Button3 = Button(self.Frame1)
        self.Button3.place(relx=0.15, rely=0.35, height=85, width=300)
        self.Button3.configure(activebackground="beige")
        self.Button3.configure(activeforeground="#000000")
        self.Button3.configure(background="#062830")
        self.Button3.configure(disabledforeground="#bfbfbf")
        self.Button3.configure(font=font14)
        self.Button3.configure(foreground="white")
        self.Button3.configure(highlightbackground="#d9d9d9")
        self.Button3.configure(highlightcolor="black")
        self.Button3.configure(pady="0")
        self.Button3.configure(text='''GUEST LIST''')
        self.Button3.configure(width=566)
        self.Button3.configure(command=click_list)

        self.Button4 = Button(self.Frame1)
        self.Button4.place(relx=0.15, rely=0.50, height=85, width=300)
        self.Button4.configure(activebackground="beige")
        self.Button4.configure(activeforeground="#000000")
        self.Button4.configure(background="#062830")
        self.Button4.configure(disabledforeground="#bfbfbf")
        self.Button4.configure(font=font14)
        self.Button4.configure(foreground="white")
        self.Button4.configure(highlightbackground="#d9d9d9")
        self.Button4.configure(highlightcolor="white")
        self.Button4.configure(pady="0")
        self.Button4.configure(text='''CHECK OUT''')
        self.Button4.configure(width=566)
        self.Button4.configure(command=click_checkout)

        self.Button5 = Button(self.Frame1)
        self.Button5.place(relx=0.15, rely=0.65, height=85, width=300)
        self.Button5.configure(activebackground="beige")
        self.Button5.configure(activeforeground="#000000")
        self.Button5.configure(background="#062830")
        self.Button5.configure(disabledforeground="#bfbfbf")
        self.Button5.configure(font=font14)
        self.Button5.configure(foreground="white")
        self.Button5.configure(highlightbackground="#d9d9d9")
        self.Button5.configure(highlightcolor="black")
        self.Button5.configure(pady="0")
        self.Button5.configure(text='''GET INFO OF ANY GUEST''')
        self.Button5.configure(width=566)
        self.Button5.configure(command=click_getinfo)

        self.Button6 = Button(self.Frame1)
        self.Button6.place(relx=0.15, rely=0.80, height=85, width=300)
        self.Button6.configure(activebackground="beige")
        self.Button6.configure(activeforeground="#000000")
        self.Button6.configure(background="#062830")
        self.Button6.configure(disabledforeground="#bfbfbf")
        self.Button6.configure(font=font14)
        self.Button6.configure(foreground="white")
        self.Button6.configure(highlightbackground="#d9d9d9")
        self.Button6.configure(highlightcolor="black")
        self.Button6.configure(pady="0")
        self.Button6.configure(text='''EXIT''')
        self.Button6.configure(width=566)
        self.Button6.configure(command=quit)
    
# Load an image
        f=Frame(root, bd=3,bg="#062830", width=510,height=500,relief=GROOVE)
        f.place(x=630,y=145)

        img=PhotoImage(file="Images/pricelist.png")
        lbl=Label(f,bg="black",width=500, height=486, image=img)
        lbl.place(x=0,y=0)  
        root.mainloop()

       
       
       
        root.mainloop()


if __name__ == '__main__':
    GUUEST=HOMESTAY_MANAGEMENT()



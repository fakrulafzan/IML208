
import os
import pickle
import sys
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
details_list=[]


def file_save():
    NAME_PRO = details_list[0]
    ADDRESS_PRO = details_list[1]
    MOBILE_NO_PRO = details_list[2]
    ROOM_NO_PRO = details_list[3]
    PRICE_PRO = details_list[4]
    f = open("homestay.dat", "ab")
    a=save(NAME_PRO,ADDRESS_PRO,MOBILE_NO_PRO,ROOM_NO_PRO,PRICE_PRO)
    pickle.dump(a,f,protocol=2)
    f.close()
    listq=[str(NAME_PRO),str(ADDRESS_PRO),str(MOBILE_NO_PRO),str(ROOM_NO_PRO),str(PRICE_PRO)]
    myVars = {'A':NAME_PRO,"B":ADDRESS_PRO,"C":MOBILE_NO_PRO,"D":ROOM_NO_PRO,"E":PRICE_PRO }

    fo=open("recipt.txt","w+")
    for h in range(0,5):
        fo.write(listq[h]+"\r\n")
    fo.close()
    call(["python", "recipt.py"])
    restart_program()






u = list()
La_Casa= (1991,1992)
La_Vouge = (1993,1994)
La_Delux= (1995,1996)
La_Villa = (1997,1998)
m = [9]
G = []
def restart_program():
    """Restarts the current program.
    Note: this function does not return. Any cleanup action (like
    saving data) must be done before calling this function."""
    python = sys.executable
    os.execl(python, python, * sys.argv)


class save:
    def __init__(self,NAME_PRO,ADDRESS_PRO,MOBILE_NO_PRO,ROOM_NO_PRO,PRICE_PRO):
        self.name=NAME_PRO
        self.address=ADDRESS_PRO
        self.mobile_no=MOBILE_NO_PRO
        self.room_no=ROOM_NO_PRO
        self.price=PRICE_PRO



class HOMESTAY_MANAGEMENT_checkin:


    def __init__(self):
        self.NAME=""
        self.ADDERESS=""
        self.MOBILE=""
        self.DAYS=0
        self.price=0
        self.room=0





        def chk_name():
            while True:

                self.k = str(self.name.get())

                a = self.k.isdigit()
                if len(self.k) != 0 and a != True:
                    self.NAME=self.k
                    self.Text1.insert(INSERT, "name has been inputed""\n")
                    break
                else:
                    self.Text1.insert(INSERT, "invalid input please input a valid name""\n")

                    break

        def chk_add():
            while True:
                self.g = str(self.addr.get())


                ak = self.g.isdigit()
                if len(self.g)!= 0 and ak!=True:
                    self.ADDERESS=self.g
                    self.Text1.insert(INSERT, "address has been inputed""\n")
                    break
                else:
                    self.Text1.insert(INSERT, "invalid input please input a valid address""\n")

                    break

        def chk_mo():
            while True:

                self.h = str(self.mobile.get())
                if self.h.isdigit() == True and len(self.h) != 0 and len(self.h) == 10:
                    self.MOBILE = self.h
                    self.Text1.insert(INSERT, "mobile number has been inputed""\n")
                    break
                else:
                    self.Text1.insert(INSERT, "invalid input please input a valid mobile number""\n")
                break

        def chk_day():
            while True:

                self.l = str(self.days.get())

                if self.l.isdigit() == True and len(self.l) != 0:
                    self.DAYS = int(self.l)
                    self.Text1.insert(INSERT, "days has been inputed""\n")
                    break
                else:
                    self.Text1.insert(INSERT, "invalid input ""\n")
                    break

        def enter(self):
            self.name = self.NAME
            self.address = self.ADDERESS
            self.mobile_no = self.MOBILE
            self.no_of_days = int(self.DAYS)

        def tor(self):

            if self.ch == 1:
                self.price = self.price + (500* self.no_of_days)
                m[0] = 1
            elif self.ch == 2:
                self.price = self.price + (1000 * self.no_of_days)
                m[0] = 2
            elif self.ch == 3:
                self.price = self.price + (1500 * self.no_of_days)
                m[0] = 3
            elif self.ch == 4:
                self.price = self.price + (1820 * self.no_of_days)
                m[0] = 4

        def payment_option(self):
            op = self.p
            if op == 1:
                self.Text1.insert(INSERT, "no discount""\n")
            elif op == 2:
                self.price = self.price - ((self.price * 10) / 100)
                self.Text1.insert(INSERT, "10% discount""\n")

        def bill(self):

            if m[0] == 1:
                a = La_Casa
            elif m[0] == 2:
                a = La_Vouge
            elif m[0] == 3:
                a = La_Delux
            elif m[0] == 4:
                a = La_Villa

            G = []
            f2 = open("homestay.dat", "rb")
            try:
                while True:
                    s = pickle.load(f2)

                    k = s.room_no
                    G.append(k)
                    continue

            except EOFError:
                pass

            for r in a:
                if r not in G:
                    self.room = r
                    break
                else:
                    continue
            self.room = r
            f2.close()

            details_list.append(self.name)
            details_list.append(self.address)
            details_list.append(self.mobile_no)
            details_list.append(self.room)
            details_list.append(self.price)




            file_save()



        def submit_clicked():
            if self.var1.get()==1 and self.var2.get()==0 and self.var3.get()==0 and self.var4.get()==0 and self.var5.get()==1 and self.var6.get()==0:
                self.ch=1
                self.p=2

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)


            elif self.var1.get() == 1 and self.var2.get() == 0 and self.var3.get() == 0 and self.var4.get() == 0 and self.var5.get() == 0 and self.var6.get() == 1:
                self.ch = 1
                self.p = 1

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)
            elif self.var1.get() == 0 and self.var2.get() == 1 and self.var3.get() == 0 and self.var4.get() == 0 and self.var5.get() == 0 and self.var6.get() == 1:
                self.ch = 2
                self.p = 1

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)
            elif self.var1.get() == 0 and self.var2.get() == 1 and self.var3.get() == 0 and self.var4.get() == 0 and self.var5.get() == 1 and self.var6.get() ==0 :
                self.ch = 2
                self.p = 2

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)
            elif self.var1.get() == 0 and self.var2.get() == 0 and self.var3.get() == 1 and self.var4.get() == 0 and self.var5.get() == 0 and self.var6.get() == 1:
                self.ch = 3
                self.p = 1

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)
            elif self.var1.get() == 0 and self.var2.get() == 0 and self.var3.get() == 1 and self.var4.get() == 0 and self.var5.get() == 1 and self.var6.get() == 0:
                self.ch = 3
                self.p = 2

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)

            elif self.var1.get() == 0 and self.var2.get() == 0 and self.var3.get() == 0 and self.var4.get() == 1 and self.var5.get() == 0 and self.var6.get() == 1:
                self.ch = 4
                self.p = 1

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)
            elif self.var1.get() == 0 and self.var2.get() == 0 and self.var3.get() == 0 and self.var4.get() == 1 and self.var5.get() == 1 and self.var6.get() == 0:
                self.ch = 4
                self.p = 2

                enter(self)
                tor(self)
                payment_option(self)
                bill(self)

            else:
                self.Text1.insert(INSERT, "invalid choice please input a valid choice""\n")





        root = Tk()


        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#062830'  # X11 color: 'white'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#ffffff' # X11 color: 'white'
        _ana1color = '#ffffff' # X11 color: 'white'
        _ana2color = '#ffffff' # X11 color: 'white'
        font10 = "-family {Courier New} -size 10 -weight normal -slant"  \
            " roman -underline 0 -overstrike 0"
        font11 = "-family {Segoe UI} -size 30 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font12 = "-family {Segoe UI} -size 18 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font13 = "-family {Segoe UI} -size 17 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font14 = "-family {Segoe UI} -size 16 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font15 = "-family {Segoe UI} -size 19 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font16 = "-family {Segoe UI} -size 15 -weight bold -slant "  \
            "roman -underline 0 -overstrike 0"
        font9 = "-family {Segoe UI} -size 9 -weight normal -slant "  \
            "roman -underline 0 -overstrike 0"

        root.geometry("1069x742")
        root.title("BOOKING MANAGEMENT")
        root.configure(background="#062830")
        root.configure(highlightbackground="#ffffff")
        root.configure(highlightcolor="black")

        self.Text1 = Text(root)
        self.Text1.place(relx=0.03, rely=0.65, relheight=0.29, relwidth=0.93)
        self.Text1.configure(background="beige")
        self.Text1.configure(font=font9)
        self.Text1.configure(foreground="black")
        self.Text1.configure(highlightbackground="#ffffff")
        self.Text1.configure(highlightcolor="black")
        self.Text1.configure(insertbackground="black")
        self.Text1.configure(selectbackground="#e6e6e6")
        self.Text1.configure(selectforeground="black")
        self.Text1.configure(width=994)
        self.Text1.configure(wrap=WORD)

        self.Frame1 = Frame(root)
        self.Frame1.place(relx=0.03, rely=0.05, relheight=0.12, relwidth=0.93)
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(borderwidth="2")
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(background="beige")
        self.Frame1.configure(highlightbackground="#ffffff")
        self.Frame1.configure(highlightcolor="black")
        self.Frame1.configure(width=995)

        self.Message1 = Message(self.Frame1)
        self.Message1.place(relx=0.250, rely=0.11, relheight=0.84, relwidth=0.5)
        self.Message1.configure(background="beige")
        self.Message1.configure(font=font11)
        self.Message1.configure(foreground="#000000")
        self.Message1.configure(highlightbackground="#ffffff")
        self.Message1.configure(highlightcolor="black")
        self.Message1.configure(text='''CUSTOMER BOOKING ''')
        self.Message1.configure(width=496)

        self.menubar = Menu(root,font=font9,bg=_bgcolor,fg=_fgcolor)
        root.configure(menu = self.menubar)



        self.Frame2 = Frame(root)
        self.Frame2.place(relx=0.03, rely=0.18, relheight=0.46, relwidth=0.93)
        self.Frame2.configure(relief=GROOVE)
        self.Frame2.configure(borderwidth="2")
        self.Frame2.configure(relief=GROOVE)
        self.Frame2.configure(background="beige")
        self.Frame2.configure(highlightbackground="#ffffff")
        self.Frame2.configure(highlightcolor="black")
        self.Frame2.configure(width=995)

        self.Label3 = Label(self.Frame2)
        self.Label3.place(relx=0.04, rely=0.03, height=47, width=289)
        self.Label3.configure(activebackground="#ffffff")
        self.Label3.configure(activeforeground="black")
        self.Label3.configure(background="beige")
        self.Label3.configure(disabledforeground="#bfbfbf")
        self.Label3.configure(font=font12)
        self.Label3.configure(foreground="#000000")
        self.Label3.configure(highlightbackground="#ffffff")
        self.Label3.configure(highlightcolor="black")
        self.Label3.configure(text='''NAME''')

        self.Label4 = Label(self.Frame2)
        self.Label4.place(relx=0.075, rely=0.29, height=47, width=329)
        self.Label4.configure(activebackground="#ffffff")
        self.Label4.configure(activeforeground="black")
        self.Label4.configure(background="beige")
        self.Label4.configure(disabledforeground="#bfbfbf")
        self.Label4.configure(font=font12)
        self.Label4.configure(foreground="#000000")
        self.Label4.configure(highlightbackground="#ffffff")
        self.Label4.configure(highlightcolor="black")
        self.Label4.configure(text='''PHONE NUMBER''')



        self.Entry3 = Entry(self.Frame2)
        self.name=StringVar()
        self.Entry3.place(relx=0.47, rely=0.05,height=34, relwidth=0.43)
        self.Entry3.configure(background="beige")
        self.Entry3.configure(disabledforeground="#bfbfbf")
        self.Entry3.configure(font=font10)
        self.Entry3.configure(foreground="#000000")
        self.Entry3.configure(highlightbackground="#ffffff")
        self.Entry3.configure(highlightcolor="black")
        self.Entry3.configure(insertbackground="black")
        self.Entry3.configure(selectbackground="#e6e6e6")
        self.Entry3.configure(selectforeground="black")
        self.Entry3.configure(textvariable=self.name)


        self.Entry4 = Entry(self.Frame2)
        self.mobile=StringVar()
        self.Entry4.place(relx=0.47, rely=0.31,height=34, relwidth=0.43)
        self.Entry4.configure(background="beige")
        self.Entry4.configure(disabledforeground="#bfbfbf")
        self.Entry4.configure(font=font10)
        self.Entry4.configure(foreground="#000000")
        self.Entry4.configure(highlightbackground="#ffffff")
        self.Entry4.configure(highlightcolor="black")
        self.Entry4.configure(insertbackground="black")
        self.Entry4.configure(selectbackground="#e6e6e6")
        self.Entry4.configure(selectforeground="black")
        self.Entry4.configure(textvariable=self.mobile)


        self.Entry5 = Entry(self.Frame2)
        self.addr = StringVar()
        self.Entry5.place(relx=0.47, rely=0.18,height=34, relwidth=0.43)
        self.Entry5.configure(background="beige")
        self.Entry5.configure(disabledforeground="#bfbfbf")
        self.Entry5.configure(font=font10)
        self.Entry5.configure(foreground="#000000")
        self.Entry5.configure(highlightbackground="#ffffff")
        self.Entry5.configure(highlightcolor="black")
        self.Entry5.configure(insertbackground="black")
        self.Entry5.configure(selectbackground="#e6e6e6")
        self.Entry5.configure(selectforeground="black")
        self.Entry5.configure(textvariable=self.addr)

        self.Label5 = Label(self.Frame2)
        self.Label5.place(relx=0.030, rely=0.16, height=47, width=334)
        self.Label5.configure(activebackground="#ffffff")
        self.Label5.configure(activeforeground="black")
        self.Label5.configure(background="beige")
        self.Label5.configure(disabledforeground="#bfbfbf")
        self.Label5.configure(font=font12)
        self.Label5.configure(foreground="#000000")
        self.Label5.configure(highlightbackground="#ffffff")
        self.Label5.configure(highlightcolor="black")
        self.Label5.configure(text='''ADDRESS''')

        self.Label6 = Label(self.Frame2)
        self.Label6.place(relx=0.32, rely=0.5, height=48, width=296)
        self.Label6.configure(activebackground="#ffffff")
        self.Label6.configure(activeforeground="black")
        self.Label6.configure(background="beige")
        self.Label6.configure(disabledforeground="#bfbfbf")
        self.Label6.configure(font=font13)
        self.Label6.configure(foreground="#000000")
        self.Label6.configure(highlightbackground="#ffffff")
        self.Label6.configure(highlightcolor="black")
        self.Label6.configure(text='''HOMESTAY TYPE''')

        self.Label7 = Label(self.Frame2)
        self.Label7.place(relx=0.3, rely=0.79, height=48, width=300)
        self.Label7.configure(activebackground="#ffffff")
        self.Label7.configure(activeforeground="black")
        self.Label7.configure(background="beige")
        self.Label7.configure(disabledforeground="#bfbfbf")
        self.Label7.configure(font=font14)
        self.Label7.configure(foreground="#000000")
        self.Label7.configure(highlightbackground="#ffffff")
        self.Label7.configure(highlightcolor="black")
        self.Label7.configure(text='''PAYMENT METHOD''')

        self.Message4 = Message(self.Frame2)
        self.Message4.place(relx=0.41, rely=0.04, relheight=0.1, relwidth=0.05)
        self.Message4.configure(background="beige")
        self.Message4.configure(font=font13)
        self.Message4.configure(foreground="#000000")
        self.Message4.configure(highlightbackground="#ffffff")
        self.Message4.configure(highlightcolor="black")
        self.Message4.configure(text=''':''')
        self.Message4.configure(width=46)

        self.Message5 = Message(self.Frame2)
        self.Message5.place(relx=0.42, rely=0.17, relheight=0.12, relwidth=0.03)
        self.Message5.configure(background="beige")
        self.Message5.configure(font=font13)
        self.Message5.configure(foreground="#000000")
        self.Message5.configure(highlightbackground="#ffffff")
        self.Message5.configure(highlightcolor="black")
        self.Message5.configure(text=''':''')
        self.Message5.configure(width=26)

        self.Message6 = Message(self.Frame2)
        self.Message6.place(relx=0.415, rely=0.3, relheight=0.09, relwidth=0.04)
        self.Message6.configure(background="beige")
        self.Message6.configure(font=font13)
        self.Message6.configure(foreground="#000000")
        self.Message6.configure(highlightbackground="#ffffff")
        self.Message6.configure(highlightcolor="black")
        self.Message6.configure(text=''':''')
        self.Message6.configure(width=36)

        self.Checkbutton1 = Checkbutton(self.Frame2)
        self.var1 = IntVar()
        self.Checkbutton1.place(relx=0.15, rely=0.60, relheight=0.17, 
                                relwidth=0.20)
        self.Checkbutton1.configure(activebackground="#ffffff")
        self.Checkbutton1.configure(activeforeground="#000000")
        self.Checkbutton1.configure(background="beige")
        self.Checkbutton1.configure(disabledforeground="#bfbfbf")
        self.Checkbutton1.configure(font=font14)
        self.Checkbutton1.configure(foreground="#000000")
        self.Checkbutton1.configure(highlightbackground="#ffffff")
        self.Checkbutton1.configure(highlightcolor="black")
        self.Checkbutton1.configure(justify=LEFT)
        self.Checkbutton1.configure(text='''LA CASA''')
        self.Checkbutton1.configure(variable=self.var1)





        self.Checkbutton2 = Checkbutton(self.Frame2)
        self.var2 = IntVar()
        self.Checkbutton2.place(relx=0.15, rely=0.72, relheight=0.11
                , relwidth=0.21)
        self.Checkbutton2.configure(activebackground="#ffffff")
        self.Checkbutton2.configure(activeforeground="#000000")
        self.Checkbutton2.configure(background="beige")
        self.Checkbutton2.configure(disabledforeground="#bfbfbf")
        self.Checkbutton2.configure(font=font13)
        self.Checkbutton2.configure(foreground="#000000")
        self.Checkbutton2.configure(highlightbackground="#ffffff")
        self.Checkbutton2.configure(highlightcolor="black")
        self.Checkbutton2.configure(justify=LEFT)
        self.Checkbutton2.configure(text='''LA VOUGE''')
        self.Checkbutton2.configure(variable=self.var2)

        self.Checkbutton3 = Checkbutton(self.Frame2)
        self.var3 = IntVar()
        self.Checkbutton3.place(relx=0.5, rely=0.62, relheight=0.11
                , relwidth=0.16)
        self.Checkbutton3.configure(activebackground="#ffffff")
        self.Checkbutton3.configure(activeforeground="#000000")
        self.Checkbutton3.configure(background="beige")
        self.Checkbutton3.configure(disabledforeground="#bfbfbf")
        self.Checkbutton3.configure(font=font13)
        self.Checkbutton3.configure(foreground="#000000")
        self.Checkbutton3.configure(highlightbackground="#ffffff")
        self.Checkbutton3.configure(highlightcolor="black")
        self.Checkbutton3.configure(justify=LEFT)
        self.Checkbutton3.configure(text='''LA DELUX''')
        self.Checkbutton3.configure(variable=self.var3)

        self.Checkbutton4 = Checkbutton(self.Frame2)
        self.var4 = IntVar()
        self.Checkbutton4.place(relx=0.5, rely=0.71, relheight=0.11
                                , relwidth=0.15)
        self.Checkbutton4.configure(activebackground="#ffffff")
        self.Checkbutton4.configure(activeforeground="#000000")
        self.Checkbutton4.configure(background="beige")
        self.Checkbutton4.configure(disabledforeground="#bfbfbf")
        self.Checkbutton4.configure(font=font13)
        self.Checkbutton4.configure(foreground="#000000")
        self.Checkbutton4.configure(highlightbackground="#ffffff")
        self.Checkbutton4.configure(highlightcolor="black")
        self.Checkbutton4.configure(justify=LEFT)
        self.Checkbutton4.configure(text='''LA VILLA''')
        self.Checkbutton4.configure(variable=self.var4)

        self.Checkbutton5 = Checkbutton(self.Frame2)
        self.var5 = IntVar()
        self.Checkbutton5.place(relx=0.485, rely=0.89, relheight=0.1
                                , relwidth=0.3)
        self.Checkbutton5.configure(activebackground="#ffffff")
        self.Checkbutton5.configure(activeforeground="#000000")
        self.Checkbutton5.configure(background="beige")
        self.Checkbutton5.configure(disabledforeground="#bfbfbf")
        self.Checkbutton5.configure(font=font16)
        self.Checkbutton5.configure(foreground="#000000")
        self.Checkbutton5.configure(highlightbackground="#ffffff")
        self.Checkbutton5.configure(highlightcolor="black")
        self.Checkbutton5.configure(justify=LEFT)
        self.Checkbutton5.configure(text='''Credit/debit card''')
        self.Checkbutton5.configure(variable=self.var5)

        self.Checkbutton6 = Checkbutton(self.Frame2)
        self.var6 = IntVar()
        self.Checkbutton6.place(relx=0.153, rely=0.89, relheight=0.1
                                , relwidth=0.15)
        self.Checkbutton6.configure(activebackground="#ffffff")
        self.Checkbutton6.configure(activeforeground="#000000")
        self.Checkbutton6.configure(background="beige")
        self.Checkbutton6.configure(disabledforeground="#bfbfbf")
        self.Checkbutton6.configure(font=font16)
        self.Checkbutton6.configure(foreground="#000000")
        self.Checkbutton6.configure(highlightbackground="#ffffff")
        self.Checkbutton6.configure(highlightcolor="black")
        self.Checkbutton6.configure(justify=LEFT)
        self.Checkbutton6.configure(text='''Cash''')
        self.Checkbutton6.configure(variable=self.var6)

        self.Message7 = Message(self.Frame2)
        self.Message7.place(relx=0.28, rely=0.46, relheight=0.11, relwidth=0.04)
        self.Message7.configure(background="beige")
        self.Message7.configure(font=font15)
        self.Message7.configure(foreground="#000000")
        self.Message7.configure(highlightbackground="#ffffff")
        self.Message7.configure(highlightcolor="black")
        self.Message7.configure(text='''-''')
        self.Message7.configure(width=41)

        self.Button1 = Button(self.Frame2)
        self.Button1.place(relx=0.91, rely=0.05, height=33, width=43)
        self.Button1.configure(activebackground="#ffffff")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#062830")
        self.Button1.configure(disabledforeground="#bfbfbf")
        self.Button1.configure(foreground="white")
        self.Button1.configure(highlightbackground="#ffffff")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''OK''')
        self.Button1.configure(command=chk_name)

        self.Button2 = Button(self.Frame2)
        self.Button2.place(relx=0.91, rely=0.18, height=33, width=43)
        self.Button2.configure(activebackground="#ffffff")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#062830")
        self.Button2.configure(disabledforeground="#bfbfbf")
        self.Button2.configure(foreground="white")
        self.Button2.configure(highlightbackground="#ffffff")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''OK''')
        self.Button2.configure(command=chk_add)

        self.Button3 = Button(self.Frame2)
        self.Button3.place(relx=0.91, rely=0.31, height=33, width=43)
        self.Button3.configure(activebackground="#ffffff")
        self.Button3.configure(activeforeground="#000000")
        self.Button3.configure(background="#062830")
        self.Button3.configure(disabledforeground="#bfbfbf")
        self.Button3.configure(foreground="white")
        self.Button3.configure(highlightbackground="#ffffff")
        self.Button3.configure(highlightcolor="black")
        self.Button3.configure(pady="0")
        self.Button3.configure(text='''OK''')
        self.Button3.configure(command=chk_mo)

        self.Button4 = Button(self.Frame2)
        self.Button4.place(relx=0.76, rely=0.66, height=83, width=156)
        self.Button4.configure(activebackground="#ffffff")
        self.Button4.configure(activeforeground="#000000")
        self.Button4.configure(background="#062830")
        self.Button4.configure(disabledforeground="#bfbfbf")
        self.Button4.configure(font=font16)
        self.Button4.configure(foreground="white")
        self.Button4.configure(highlightbackground="#ffffff")
        self.Button4.configure(highlightcolor="black")
        self.Button4.configure(pady="0")
        self.Button4.configure(text='''SUBMIT''')
        self.Button4.configure(command=submit_clicked)

        self.Label1 = Label(self.Frame2)
        self.Label1.place(relx=0.112, rely=0.43, height=44, width=260)
        self.Label1.configure(background="beige")
        self.Label1.configure(disabledforeground="#bfbfbf")
        self.Label1.configure(font=font13)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(text='''NUMBER OF DAYS''')


        self.Entry1 = Entry(self.Frame2)
        self.days=StringVar()
        self.Entry1.place(relx=0.47, rely=0.43, height=34, relwidth=0.43)
        self.Entry1.configure(background="beige")
        self.Entry1.configure(disabledforeground="#bfbfbf")
        self.Entry1.configure(font=font10)
        self.Entry1.configure(foreground="#000000")
        self.Entry1.configure(insertbackground="black")
        self.Entry1.configure(width=424)
        self.Entry1.configure(textvariable=self.days)

        self.Message8 = Message(self.Frame2)
        self.Message8.place(relx=0.42, rely=0.41, relheight=0.11, relwidth=0.03)
        self.Message8.configure(background="beige")
        self.Message8.configure(font=font13)
        self.Message8.configure(foreground="#000000")
        self.Message8.configure(highlightbackground="#ffffff")
        self.Message8.configure(highlightcolor="black")
        self.Message8.configure(text=''':''')
        self.Message8.configure(width=26)

        self.Button5 = Button(self.Frame2)
        self.Button5.place(relx=0.91, rely=0.43, height=33, width=43)
        self.Button5.configure(activebackground="#ffffff")
        self.Button5.configure(activeforeground="#000000")
        self.Button5.configure(background="#062830")
        self.Button5.configure(disabledforeground="#bfbfbf")
        self.Button5.configure(foreground="white")
        self.Button5.configure(highlightbackground="#ffffff")
        self.Button5.configure(highlightcolor="black")
        self.Button5.configure(pady="0")
        self.Button5.configure(text='''OK''')
        self.Button5.configure(command=chk_day)


        root.mainloop()


if __name__ == '__main__':
    hotel=HOMESTAY_MANAGEMENT_checkin()







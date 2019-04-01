import xlrd
import xlwt
from xlutils.copy import *
from tkinter import *
from tkinter import messagebox
from PIL import ImageTk, Image
import time

global root
root=Tk()
root.geometry('1200x800')
root.title('ATM')
r=Frame(root)
r.grid(row=0,column=0,sticky='news')
Label(r,text='ATM',font=('Aerial 70 bold')).grid(row=0,column=0)
Label(r,text='').grid(row=1,column=0)
Label(r,text='').grid(row=2,column=0)
Label(r,text='').grid(row=3,column=0)
Label(r,text='Enter username',font=('Chiller 40')).grid(row=4,column=0)
Label(r,text='').grid(row=4,column=1)
uname=Entry(r)
uname.grid(row=4,column=2)
Label(r,text='').grid(row=5,column=0)
Label(r,text='Enter Pin',font=('Chiller 40')).grid(row=6,column=0)
Label(r,text='').grid(row=6,column=1)
pwd=Entry(r)
pwd.grid(row=6,column=2)
Label(r,text='').grid(row=7,column=0)
Label(r,text='').grid(row=8,column=0)
Label(r,text='').grid(row=9,column=0)

def chkBal(user):
    fr=Frame(root)
    fr.grid(row=0,column=0,sticky='news')
    fr.tkraise()
    Label(fr,text='ATM',font=('Aerial 70 bold')).grid(row=0,column=0)
    Label(fr,text='').grid(row=1,column=0)
    Label(fr,text='').grid(row=2,column=0)
    Label(fr,text='Your balance is '+ str(ds.cell_value(user,2)),font=('Chiller 50')).grid(row=3,column=0)

def okpin(p1,p2,user):
    if(p1.get()==p2.get()):
        dwc=copy(dw)
        dwr=dwc.get_sheet(0)
        dwr.write(user,1,p1.get())
        dwc.save('user_data.xls')
        messagebox.showinfo('Congrats!!','Your pin is changed succesfully')
    else:
        messagebox.showinfo('ERROR','Both fields does not match')
        
def chngPin(user):
    fr=Frame(root)
    fr.grid(row=0,column=0,sticky='news')
    fr.tkraise()
    Label(fr,text='ATM',font=('Aerial 70 bold')).grid(row=0,column=0)
    Label(fr,text='').grid(row=1,column=0)
    Label(fr,text='').grid(row=2,column=0)
    Label(fr,text='New Pin',font=('Chiller 50')).grid(row=3,column=0)
    Label(fr,text='').grid(row=3,column=1)
    p1=Entry(fr)
    p1.grid(row=3,column=2)
    Label(fr,text='').grid(row=4,column=0)
    Label(fr,text='Confirm Pin',font=('Chiller 50')).grid(row=5,column=0)
    Label(fr,text='').grid(row=5,column=1)
    p2=Entry(fr)
    p2.grid(row=5,column=2)
    Label(fr,text='').grid(row=6,column=0)
    Label(fr,text='').grid(row=7,column=0)
    Button(fr,text='OK',command=lambda: okpin(p1,p2,user)).grid(row=7,column=1)

def okdep(p1,user):
    dwc=copy(dw)
    dwr=dwc.get_sheet(0)
    bal=int(ds.cell_value(user,2))+int(p1.get())
    dwr.write(user,2,bal)
    dwc.save('user_data.xls')
    messagebox.showinfo('Amount deposited successfully','Your balance is '+ str(bal))
    
def deposit(user):
    fr=Frame(root)
    fr.grid(row=0,column=0,sticky='news')
    fr.tkraise()
    Label(fr,text='ATM',font=('Aerial 70 bold')).grid(row=0,column=0)
    Label(fr,text='').grid(row=1,column=0)
    Label(fr,text='').grid(row=2,column=0)
    Label(fr,text='Enter amount to be deposited: ',font=('Chiller 50')).grid(row=3,column=0)
    Label(fr,text='').grid(row=3,column=1)
    p1=Entry(fr)
    p1.grid(row=3,column=2)
    Label(fr,text='').grid(row=4,column=0)
    Button(fr,text='OK',command=lambda: okdep(p1,user)).grid(row=7,column=1)

def okwit(p1,user):
    dwc=copy(dw)
    dwr=dwc.get_sheet(0)
    bal=int(ds.cell_value(user,2))-int(p1.get())
    if(bal>=3000):
        dwr.write(user,2,bal)
        dwc.save('user_data.xls')
        messagebox.showinfo('Amount withdrawn successfully','Your balance is '+ str(bal))
    else:
        messagebox.showinfo('Amount withdraw failed','Your balance is below required amount')
        
def withdraw(user):
    fr=Frame(root)
    fr.grid(row=0,column=0,sticky='news')
    fr.tkraise()
    Label(fr,text='ATM',font=('Aerial 70 bold')).grid(row=0,column=0)
    Label(fr,text='').grid(row=1,column=0)
    Label(fr,text='').grid(row=2,column=0)
    Label(fr,text='Enter amount to be withdrawn: ',font=('Chiller 50')).grid(row=3,column=0)
    Label(fr,text='').grid(row=3,column=1)
    p1=Entry(fr)
    p1.grid(row=3,column=2)
    Label(fr,text='').grid(row=4,column=0)
    Button(fr,text='OK',command=lambda: okwit(p1,user)).grid(row=7,column=1)

def oktrf(p1,p2,user):
    dwc=copy(dw)
    dwr=dwc.get_sheet(0)
    bal=int(ds.cell_value(user,2))-int(p2.get())
    if(bal>=3000):
        dnr=ds.nrows
        for k in range (1,dnr):
            a=ds.cell_value(k,0)
            if(a==p1.get()):
                user2=k
                break
            if(k==dnr-1):
                messagebox.showinfo('Amount transfer failed','Reciever does not exist')
                return 
        dwr.write(user,2,bal)
        dwr.write(user2,2,int(ds.cell_value(user2,2))+int(p2.get()))
        dwc.save('user_data.xls')
        messagebox.showinfo('Amount transfered successfully','Your balance is '+ str(bal))
    else:
        messagebox.showinfo('Amount transfer failed','Your balance is below required amount')
        
def transfer(user):
    fr=Frame(root)
    fr.grid(row=0,column=0,sticky='news')
    fr.tkraise()
    Label(fr,text='ATM',font=('Aerial 70 bold')).grid(row=0,column=0)
    Label(fr,text='').grid(row=1,column=0)
    Label(fr,text='').grid(row=2,column=0)
    Label(fr,text='Enter username of reciever',font=('Chiller 50')).grid(row=3,column=0)
    Label(fr,text='').grid(row=3,column=1)
    p1=Entry(fr)
    p1.grid(row=3,column=2)
    Label(fr,text='').grid(row=4,column=0)
    Label(fr,text='Enter amount to be transfered',font=('Chiller 50')).grid(row=5,column=0)
    Label(fr,text='').grid(row=5,column=1)
    p2=Entry(fr)
    p2.grid(row=5,column=2)
    Label(fr,text='').grid(row=6,column=0)
    Label(fr,text='').grid(row=7,column=0)
    Button(fr,text='OK',command=lambda: oktrf(p1,p2,user)).grid(row=7,column=1)

def choose():
    global dw,ds
    m1=uname.get()
    m2=pwd.get()
    dw = xlrd.open_workbook('user_data.xls')
    ds=dw.sheet_by_index(0)
    dnr1=ds.nrows
    for i in range(1,dnr1):
        a=ds.cell_value(i,0)
        b=ds.cell_value(i,1)
        b=str(int(b))
        l=0
        for j in b:
            l +=1
        while(l!=4):
            b='0'+b
            l=0
            for j in b:
                    l +=1
        if(a==m1 and b==m2):
            messagebox.showinfo('LOGIN SUCCESSFUL','WELCOME TO ATM')
            fr=Frame(root)
            fr.grid(row=0,column=0,sticky='news')
            fr.tkraise()
            Label(fr,text='ATM',font=('Aerial 70 bold')).grid(row=1,column=1)
            Label(fr,text='').grid(row=2,column=1)
            Label(fr,text='').grid(row=3,column=1)
            Label(fr,text='What do you want to do???',font=('Chiller 50')).grid(row=4,column=1)
            Button(fr,text='Check Balance',command=lambda: chkBal(i)).grid(row=5,column=1)
            Label(fr,text='').grid(row=5,column=2)
            Button(fr,text='Change Pin',command=lambda: chngPin(i)).grid(row=5,column=3)
            Label(fr,text='').grid(row=5,column=4)
            Button(fr,text='Deposit Money',command=lambda: deposit(i)).grid(row=5,column=5)
            Label(fr,text='').grid(row=6,column=1)
            Label(fr,text='').grid(row=7,column=1)
            Button(fr,text='Withdraw Money',command=lambda: withdraw(i)).grid(row=7,column=2)
            Label(fr,text='').grid(row=7,column=3)
            Button(fr,text='Transfer Amount',command=lambda: transfer(i)).grid(row=7,column=4)
            break
        if(i==dnr1-1):
            messagebox.showinfo('ERROR','Username or Pin incorrect')
    
Button(r,text='OK',command=choose).grid(row=9,column=1)
r.tkraise()
r.mainloop()



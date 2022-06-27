from tkinter import messagebox, colorchooser
from tkinter.messagebox import showerror
import time  
import threading
from threading import Thread
import win32com.client as win32
import tkinter as tk
from tkinter import ttk
import datetime
import matplotlib.pyplot as plt
import os
import pathlib


flag = True
var_pos= {}
var_neg= {}
var_no_resp = {}

class Sinus(threading.Thread):
    def __init__(self, start, end, marshalled_iu):
        self.stop_event = threading.Event()
        self.marshalled = marshalled_iu
        super().__init__(target=self.sin, args=( fesantionare, fsemnal, amplitudine, self.stop_event,))
    
    def sin(self, fesantionare, fsemnal, amplitudine):
        app.list.delete(0,'end')    
        win32.pythoncom.CoInitialize()
        app.writ('Starting sinus...')
        fs = int(fesantionare) # frecventa esantionare
        A = amplitudine
        f0 = int(fsemnal) #freventa semnal
        tstep = 1/fs # interval timp esantion
        N = int(fs/f0)
        t = np.linspace(0, (N-1)*tstep, N)
        fstep = fs/N
        f=np.linspace(0, (N-1)*fstep, N)
        y = A*np.sin(2*np.pi*f0*t)
        app.list.delete(0,'end')
        app.writ('Finished sending requests.')
        win32.pythoncom.CoUninitialize()
        app.button11['state'] = tk.NORMAL
        save_report()
        app.save_excel()
        return y
        
    def cancel(self):
        self.stop_event.set()

        
class Ramp(threading.Thread):
    def __init__(self, i, lista_mut, marshalled_ie):
        self.stop_event = threading.Event()
        self.marshalled = marshalled_ie
        super().__init__(target=self.ramp_sig, args=(self.stop_event, fesantionare, fsemnal, amplitudine,))

    def ramp_sig(self, fesantionare, fsemnal, amplitudine):
        
        app.list.delete(0,'end')    
        win32.pythoncom.CoInitialize()
        self.ie = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(self.marshalled, pythoncom.IID_IDispatch))
        index = 2
        app.list.insert(0,'Sending requests...')
        fs = int(fesantionare) # frecventa esantionare
        A = amplitudine
        f0 = int(fsemnal) #freventa semnal
        tstep = 1/fs # interval timp esantion
        N = int(fs/f0)
        t = np.linspace(0, (N-1)*tstep, N)
        fstep = fs/N
        f=np.linspace(0, (N-1)*fstep, N)
        z = np.abs(A*signal.sawtooth(2*np.pi*f0*t))


    
        app.list.delete(0,'end')    
        app.writ('Finished sending requests.')
        win32.pythoncom.CoUninitialize()
        save_report()
        app.save_excel()
        app.button30['state'] = tk.NORMAL
        app.button2['state'] = tk.NORMAL
        return z
    
    def cancel(self):
        self.stop_event.set()

class Square(threading.Thread):
    def __init__(self, incercari, marshalled_iuc):
        self.stop_event = threading.Event()
        self.marshalled = marshalled_iuc
        var_pos['3']= 0
        var_neg['3']= 0
        var_no_resp['3']=0
        super().__init__(target=self.drept, args=(self.stop_event, fesantionare, fsemnal, amplitudine,))

    def drept(self, fesantionare, fsemnal, amplitudine):
        app.list.delete(0,'end')
        
        win32.pythoncom.CoInitialize()
        self.iud = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(self.marshalled, pythoncom.IID_IDispatch))
        flag = 0
        
        app.list.insert(0,'Starting seed entropy test...')
        fs = int(fesantionare) # frecventa esantionare
        A = amplitudine
        f0 = int(fsemnal) #freventa semnal
        tstep = 1/fs # interval timp esantion
        N = int(fs/f0)
        t = np.linspace(0, (N-1)*tstep, N)
        fstep = fs/N
        f=np.linspace(0, (N-1)*fstep, N)
        x = A*signal.square(2*np.pi*f0*t)
        win32.pythoncom.CoUninitialize()
        return x     
                              
    def cancel(self):
        self.stop_event.set()
        



class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Licenta')
        self.geometry('600x600')
        self.resizable(0, 0)
        self.style=ttk.Style(self)
        self.configure(bg='LightBlue1')
        self.style.theme_use("clam")
        self.i = {'contor':0}
        self.counter = 0
        self.vector_esantioane = []       
        self.create_header_frame()
        self.sinus_thread = None
        self.ramp_thread = None
        self.square_thread = None
        self.path = None
        self.fullpath = None
        
    def handle_sinus(self):
        if self.fesantionare.get() != 0 and self.amplitudine.get() != 0 and self.frecventa.get() !=0:
            marshalled_ief = win32.pythoncom.CoMarshalInterThreadInterfaceInStream(win32.pythoncom.IID_IDispatch, win32.spi)
            self.button30['state'] = tk.DISABLED
            self.button2['state'] = tk.DISABLED
            self.sinus_thread = Sinus(self.fesantionare.get(), self.amplitudine.get(), self.freventa.get(), marshalled_ief)
            self.sinus_thread.start()
        else:
            showerror(title='Error',
                      message='Please complete empty fields.')

    def handle_ramp(self):
        if self.lista_mutatii_f and self.serviciu_flooding != "Select service":
            marshalled_flooding = win32.pythoncom.CoMarshalInterThreadInterfaceInStream(win32.pythoncom.IID_IDispatch, win32.sys)
            self.button40['state'] = tk.DISABLED
            self.button41['state'] = tk.DISABLED
            self.ramp_thread = Ramp(self.fesantionare.get(), self.amplitudine.get(), self.freventa.get(), marshalled_flooding)
            self.ramp_thread.start()
        else:
            showerror(title='Error',
                      message='Please complete empty fields.')
    
    def start_square(self):
        
        if not self.start_flooding.get() or not self.end_flooding.get():
            showerror(title="Warning!", message="Invalid input! String is empty.")
            return

        if end_int < start_int:
            showerror(title="Warning!", message="End value must be greater than start value.")
            return
        marshalled_flood = win32.pythoncom.CoMarshalInterThreadInterfaceInStream(win32.pythoncom.IID_IDispatch)
        self.square_thread = Square(self.fesantionare.get(), self.amplitudine.get(), self.freventa.get(), marshalled_flood)
        self.square_thread.start()
        self.button45['state'] = tk.DISABLED

    
    def stop_commands(self):
        global flag
        if self.sinus_thread is not None:
            self.sinus_thread.cancel()
            self.sinus_thread = None
            self.button30['state'] = tk.NORMAL
            self.button2['state'] = tk.NORMAL
            flag = False
            self.save_excel()
            save_report()

        if self.square_thread is not None:
            self.square_thread.cancel()
            self.square_thread = None
            self.button11['state'] = tk.NORMAL
            self.save_excel()
            save_report()

        if self.ramp_thread is not None:
            self.ramp_thread.cancel()
            self.ramp_thread = None
            self.button40['state'] = tk.NORMAL
            self.button41['state'] = tk.NORMAL
            flag = False
            self.save_excel()
            save_report()
            

    def writ(self, text):
        if self.counter < 15:
            self.counter += 1
            self.list.insert(self.counter,text)
        else:
            self.list.delete(0)
            self.list.insert(self.counter,text)


        
    def create_header_frame(self):


          
        self.photo1 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo1.png')
        self.photo2 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo2.png')
        self.photo3 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo3.png')
        self.photo4 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo4.png')
        self.photo5 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo5.png')
        self.photo6 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo6.png')
        self.photo7 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo7.png')

        
        self.s = ttk.Style(self)
        self.s.configure("My.TFrame", background ="LightBlue1")
        self.s.configure('TButton', background='LightBlue1')
        self.s.configure('TButton', foreground='black')
        self.s.configure("TMenubutton", background="LightBlue1")
        self.header = ttk.Notebook(self)
        
        self.tab1 = ttk.Frame(self.header, style = "My.TFrame")
        self.header.add(self.tab1, text='Sinus')

        self.tab2 = ttk.Frame(self.header, style = "My.TFrame")
        self.header.add(self.tab2, text='Ramp')

        self.tab3 = ttk.Frame(self.header, style = "My.TFrame")
        self.header.add(self.tab3, text='Square')


        self.header.pack(expand=1, fill="both")

        self.amplitudine = tk.IntVar()
        self.frecventa = tk.IntVar()
        self.fesantionare = tk.IntVar()

        self.list = tk.Listbox(self.header, height=15, width=50, selectmode=tk.EXTENDED, bg='black', fg='grey70')
        self.list.pack(side=tk.BOTTOM, fill=tk.BOTH)

        ########################        Sinus       ######################################

        self.label1 = ttk.Label(self.tab1, text="F.Es: ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label1.place(x=200, y=60)
        self.entry1=ttk.Entry(self.tab1,font=(30), textvariable=self.fesantionare)
        self.entry1.place(x=200, y=80)

        self.label2 = ttk.Label(self.tab1, text="Frecventa : ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label2.place(x=200, y=110)
        self.entry2=ttk.Entry(self.tab1,font=(30), textvariable=self.frecventa)
        self.entry2.place(x=200, y=130)

        self.label22 = ttk.Label(self.tab1, text="Amplitudine : ",background='LightBlue1', width= 12, font=("arial",10,"bold"))
        self.label22.place(x=200, y=160)
        self.entry22=ttk.Entry(self.tab1,font=(30), textvariable=self.amplitudine)
        self.entry22.place(x=200, y=180)
        
        self.button11=tk.Button(self.tab1, image=self.photo1,bg='#BFEFFF', bd=0,width = 140, command=self.handle_sinus)
        self.button11.place(x=200, y=220)

        ###########################################################################################

        
        ########################        Ramp        ##################################
        
        self.label33 = ttk.Label(self.tab2, text="F.Es: ",background='LightBlue1', width= 5, font=("arial",10,"bold"))
        self.label33.place(x=200, y=60)
        self.entry3=ttk.Entry(self.tab2,font=(30), textvariable=self.fesantionare)
        self.entry3.place(x=200, y=80)

        self.labe34 = ttk.Label(self.tab2, text="Freventa : ",background='LightBlue1', width= 5, font=("arial",10,"bold"))
        self.labe34.place(x=200, y=120)
        self.entry4=ttk.Entry(self.tab2,font=(30), textvariable=self.frecventa)
        self.entry4.place(x=200, y=140)

        self.labe35 = ttk.Label(self.tab2, text="Amplitudine : ",background='LightBlue1', width= 5, font=("arial",10,"bold"))
        self.labe35.place(x=200, y=120)
        self.entry5=ttk.Entry(self.tab2,font=(30), textvariable=self.amplitudine)
        self.entry5.place(x=200, y=160)
    
        
        self.button30=tk.Button(self.tab2, image=self.photo6,bg='#BFEFFF', bd=0,width = 140)
        self.button30['command'] = self.handle_ramp
        self.button30.place(x=265, y=200)


        ###########################################################################################


        ###################         Square       #######################################
        
        self.labe44 = ttk.Label(self.tab3, text="F.Es: ",background='LightBlue1', width= 5, font=("arial",10,"bold"))
        self.labe44.place(x=200, y=60)
        self.entry6=ttk.Entry(self.tab3,font=(30), textvariable=self.fesantionare)
        self.entry6.place(x=200, y=80)

        self.labe45 = ttk.Label(self.tab3, text="Freventa : ",background='LightBlue1', width= 5, font=("arial",10,"bold"))
        self.labe45.place(x=200, y=120)
        self.entry7=ttk.Entry(self.tab3,font=(30), textvariable=self.frecventa)
        self.entry7.place(x=200, y=140)

        self.labe46 = ttk.Label(self.tab3, text="Amplitudine : ",background='LightBlue1', width= 5, font=("arial",10,"bold"))
        self.labe46.place(x=200, y=120)
        self.entry8=ttk.Entry(self.tab3,font=(30), textvariable=self.amplitudine)
        self.entry8.place(x=200, y=160)
        
        self.button20=tk.Button(self.tab3, image=self.photo7,bg='#BFEFFF', bd=0,width = 140 ,command=self.start_square)
        self.button20.place(x=200, y=160)


        ###########################################################################################

        ###########################################################################################

        
        self.button4=tk.Button(self.header, image=self.photo2, bg='#BFEFFF', bd=0,width = 160)
        self.button4['command'] = self.stop_commands
        self.button4.place(x=60, y=320)
        
        
        self.button5=tk.Button(self.header, image=self.photo4, bg='#BFEFFF', bd=0,width = 160)
        self.button5['command']= self.open_excel
        self.button5.place(x=220, y=320)
    def open_excel(self):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.ScreenUpdating = True
        excel.DisplayAlerts = False
        excel.EnableEvents = True
        wa = excel.Workbooks.Open(self.fullpath)
        

    def save_excel(self):
        now = datetime.datetime.now()
        
        
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        newpath = desktop + r'\Test Results'
        if not os.path.exists(newpath):
            os.makedirs(newpath)
        self.path = f"\Test_results_{now.day}_{now.month}_{now.year}_{now.hour}_{now.minute}.xlsx"
        self.fullpath = newpath + self.path
        wb.save(self.fullpath)


if __name__ == "__main__":
    app = Application()
    app.mainloop()

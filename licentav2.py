from tkinter import messagebox, colorchooser
from tkinter.messagebox import showerror
import time  
import threading
from threading import Thread
#import win32com.client as win32
import tkinter as tk
from tkinter import ttk
import datetime
import matplotlib.pyplot as plt
import os
import numpy as np
import pathlib
from drawnow import *
from scipy import signal
import ADS1256
import DAC8532
import RPi.GPIO as GPIO
flag = True
vector = []
plt.switch_backend("TkAgg")
class Sinus(threading.Thread):
    def __init__(self, fesantionare, fsemnal, amplitudine):
        self.stop_event = threading.Event()
        self.amp = amplitudine
        #self.marshalled = marshalled_iu
        super().__init__(target=self.sin, args=( fesantionare, fsemnal, amplitudine, self.stop_event,))
        
    def makefig(self):
        plt.ylim(-8,8)
        plt.xlim(0,50)
        plt.title("graphic")
        plt.grid(True)
        plt.ylabel("RandomNumber")
        plt.plot(vector, color='red', marker='o', markerfacecolor='blue', label='numere')
        plt.legend(loc='upper left')
        
    def sin(self, fesantionare, fsemnal, amplitudine, stop_ev):
        global vector
        count=0
            
        #win32.pythoncom.CoInitialize()
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
        ADC = ADS1256.ADS1256()
        DAC = DAC8532.DAC8532()
        ADC.ADS1256_init()
        while app.button1['state']==tk.DISABLED:
            app.list.delete(0,'end')
            ADC_Value = ADC.ADS1256_GetAll()
            app.list.insert (0,"0 ADC = %lf"%(ADC_Value[0]*5.0/0x7fffff))
            app.list.insert (1,"1 ADC = %lf"%(ADC_Value[1]*5.0/0x7fffff))
            app.list.insert (2,"2 ADC = %lf"%(ADC_Value[2]*5.0/0x7fffff))
            app.list.insert (3,"3 ADC = %lf"%(ADC_Value[3]*5.0/0x7fffff))
            app.list.insert (4,"4 ADC = %lf"%(ADC_Value[4]*5.0/0x7fffff))
            app.list.insert (5,"5 ADC = %lf"%(ADC_Value[5]*5.0/0x7fffff))
            app.list.insert (6,"6 ADC = %lf"%(ADC_Value[6]*5.0/0x7fffff))
            app.list.insert (7,"7 ADC = %lf"%(ADC_Value[7]*5.0/0x7fffff))
            for i in y:
                DAC.DAC8532_Out_Voltage(DAC8532.channel_A, 3.3 * i / 33)
                DAC.DAC8532_Out_Voltage(DAC8532.channel_B, 3.3 - 3.3 * i / 33)
                
            for i in y:
                DAC.DAC8532_Out_Voltage(DAC8532.channel_B, 3.3 * i / 33)
                DAC.DAC8532_Out_Voltage(DAC8532.channel_A, 3.3 - 3.3 * i / 33)
            vector.append(y[count])    
            drawnow(self.makefig)
            #plt.pause()
            count = count+1
            #if len(y) >= 50:
             #   if count>=50:
              #      vector.pop(0)
            if count>=len(y):
                vector.pop(0)
                count=0
            
            
                
        app.writ('Finished sending requests.')
        app.writ(y)
        app.button1['state'] = tk.NORMAL
        return y
    
    
    

    def cancel(self):
        self.stop_event.set()

        
class Ramp(threading.Thread):
    def __init__(self, fesantionare, fsemnal, amplitudine):
        self.stop_event = threading.Event()
        #self.marshalled = marshalled_ie
        super().__init__(target=self.ramp_sig, args=(self.stop_event, fesantionare, fsemnal, amplitudine,))
    def makefig(self):
        plt.ylim(-8,8)
        plt.xlim(0,50)
        plt.title("graphic")
        plt.grid(True)
        plt.ylabel("RandomNumber")
        plt.plot(vector, color='red', marker='o', markerfacecolor='blue', label='numere')
        plt.legend(loc='upper left')
        
    def ramp_sig(self,stop_ev, fesantionare, fsemnal, amplitudine):
        count=0
        app.list.delete(0,'end')    
        #win32.pythoncom.CoInitialize()
        #self.ie = win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(self.marshalled, pythoncom.IID_IDispatch))
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
        ADC = ADS1256.ADS1256()
        DAC = DAC8532.DAC8532()
        ADC.ADS1256_init()
        while app.button2['state']==tk.DISABLED:
            app.list.delete(0,'end')
            ADC_Value = ADC.ADS1256_GetAll()
            app.list.insert (0,"0 ADC = %lf"%(ADC_Value[0]*5.0/0x7fffff))
            app.list.insert (1,"1 ADC = %lf"%(ADC_Value[1]*5.0/0x7fffff))
            app.list.insert (2,"2 ADC = %lf"%(ADC_Value[2]*5.0/0x7fffff))
            app.list.insert (3,"3 ADC = %lf"%(ADC_Value[3]*5.0/0x7fffff))
            app.list.insert (4,"4 ADC = %lf"%(ADC_Value[4]*5.0/0x7fffff))
            app.list.insert (5,"5 ADC = %lf"%(ADC_Value[5]*5.0/0x7fffff))
            app.list.insert (6,"6 ADC = %lf"%(ADC_Value[6]*5.0/0x7fffff))
            app.list.insert (7,"7 ADC = %lf"%(ADC_Value[7]*5.0/0x7fffff))
            for i in z:
                DAC.DAC8532_Out_Voltage(DAC8532.channel_A, 3.3 * i / 33)
                DAC.DAC8532_Out_Voltage(DAC8532.channel_B, 3.3 - 3.3 * i / 33)
                
            for i in z:
                DAC.DAC8532_Out_Voltage(DAC8532.channel_B, 3.3 * i / 33)
                DAC.DAC8532_Out_Voltage(DAC8532.channel_A, 3.3 - 3.3 * i / 33)
            vector.append(z[count])    
            drawnow(self.makefig)
            #plt.pause()
            count = count+1
            #if len(z) >= 50:
             #   if count>=50:
              #      vector.pop(0)
            if count>=len(z):
                vector.pop(0)
                count=0
        app.writ('Finished sending requests.')
        app.writ(z)
        app.button2['state'] = tk.NORMAL
        return z
    
    def cancel(self):
        self.stop_event.set()

class Square(threading.Thread):
    def __init__(self, fesantionare, fsemnal, amplitudine):
        self.stop_event = threading.Event()
       # self.marshalled = marshalled_iuc
        super().__init__(target=self.drept, args=(self.stop_event, fesantionare, fsemnal, amplitudine,))
    def makefig(self):
        plt.ylim(-8,8)
        plt.xlim(0,50)
        plt.title("graphic")
        plt.grid(True)
        plt.ylabel("RandomNumber")
        plt.plot(vector, color='red', marker='o', markerfacecolor='blue', label='numere')
        plt.legend(loc='upper left')
        
    def drept(self,stop_ev, fesantionare, fsemnal, amplitudine):
        app.list.delete(0,'end')
        count=0
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
        ADC = ADS1256.ADS1256()
        DAC = DAC8532.DAC8532()
        ADC.ADS1256_init()
        while app.button3['state'] == tk.DISABLED:
            app.list.delete(0,'end')
            ADC_Value = ADC.ADS1256_GetAll()
            app.list.insert (0,"0 ADC = %lf"%(ADC_Value[0]*5.0/0x7fffff))
            app.list.insert (1,"1 ADC = %lf"%(ADC_Value[1]*5.0/0x7fffff))
            app.list.insert (2,"2 ADC = %lf"%(ADC_Value[2]*5.0/0x7fffff))
            app.list.insert (3,"3 ADC = %lf"%(ADC_Value[3]*5.0/0x7fffff))
            app.list.insert (4,"4 ADC = %lf"%(ADC_Value[4]*5.0/0x7fffff))
            app.list.insert (5,"5 ADC = %lf"%(ADC_Value[5]*5.0/0x7fffff))
            app.list.insert (6,"6 ADC = %lf"%(ADC_Value[6]*5.0/0x7fffff))
            app.list.insert (7,"7 ADC = %lf"%(ADC_Value[7]*5.0/0x7fffff))
            for i in x:
                DAC.DAC8532_Out_Voltage(DAC8532.channel_A, 3.3 * i / 33)
                DAC.DAC8532_Out_Voltage(DAC8532.channel_B, 3.3 - 3.3 * i / 33)
                
            for i in x:
                DAC.DAC8532_Out_Voltage(DAC8532.channel_B, 3.3 * i / 33)
                DAC.DAC8532_Out_Voltage(DAC8532.channel_A, 3.3 - 3.3 * i / 33)
            vector.append(x[count])    
            drawnow(self.makefig)
            #plt.pause()
            count = count+1
            #if len(x) >= 50:
             #   if count>=50:
              #      vector.pop(0)
            if count>=len(x):
                vector.pop(0)
                count=0
        app.writ('Finished sending requests.')
        app.writ(x)
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
            #marshalled_ief = win32.pythoncom.CoMarshalInterThreadInterfaceInStream(win32.pythoncom.IID_IDispatch)
            #self.button30['state'] = tk.DISABLED
            self.button1['state'] = tk.DISABLED
            self.sinus_thread = Sinus(self.fesantionare.get(), self.frecventa.get(), self.amplitudine.get())# marshalled_ief)
            self.sinus_thread.setDaemon(True)
            self.sinus_thread.start()
        else:
            showerror(title='Error',
                      message='Please complete empty fields.')

    def handle_ramp(self):
        if self.fesantionare.get() != 0 and self.amplitudine.get() != 0 and self.frecventa.get() !=0:
           # marshalled_flooding = win32.pythoncom.CoMarshalInterThreadInterfaceInStream(win32.pythoncom.IID_IDispatch, win32.sys)
            self.button2['state'] = tk.DISABLED
            self.ramp_thread = Ramp(self.fesantionare.get(),self.frecventa.get() ,self.amplitudine.get() )#, marshalled_flooding)
            
            self.ramp_thread.start()
        else:
            showerror(title='Error',
                      message='Please complete empty fields.')
    
    def start_square(self):
        
        if self.fesantionare.get() != 0 and self.amplitudine.get() != 0 and self.frecventa.get() !=0:

           # marshalled_flood = win32.pythoncom.CoMarshalInterThreadInterfaceInStream(win32.pythoncom.IID_IDispatch)
            self.square_thread = Square(self.fesantionare.get(), self.frecventa.get(),self.amplitudine.get())#, marshalled_flood)
            self.square_thread.start()
            self.button3['state'] = tk.DISABLED
        else:
            showerror(title='Error',
                      message='Please complete empty fields.')

    
    def stop_commands(self):
        global flag
        if self.sinus_thread is not None:
            self.sinus_thread.cancel()
            self.sinus_thread = None
            self.button1['state'] = tk.NORMAL
            plt.close()
            #self.button2['state'] = tk.NORMAL
            flag = False
            #self.save_excel()
            #save_report()

        if self.square_thread is not None:
            self.square_thread.cancel()
            self.square_thread = None
            self.button3['state'] = tk.NORMAL
            #self.save_excel()
            #save_report()

        if self.ramp_thread is not None:
            self.ramp_thread.cancel()
            self.ramp_thread = None
            self.button2['state'] = tk.NORMAL
            flag = False
            #self.save_excel()
            #save_report()
            

    def writ(self, text):
        if self.counter < 15:
            self.counter += 1
            self.list.insert(self.counter,text)
        else:
            self.list.delete(0)
            self.list.insert(self.counter,text)


        
    def create_header_frame(self):


          
        #self.photo1 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo1.png')
        #self.photo2 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo2.png')
        #self.photo4 = tk.PhotoImage(file = fr'{pathlib.Path().resolve()}\photo4.png')


        
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

        self.label0 = ttk.Label(self.tab1, text="F.Es: ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label0.place(x=200, y=60)
        self.entry0=ttk.Entry(self.tab1, font=(30), textvariable=self.fesantionare)
        self.entry0.place(x=200, y=80)

        self.label1 = ttk.Label(self.tab1, text="Frecventa : ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label1.place(x=200, y=110)
        self.entry1=ttk.Entry(self.tab1,font=(30), textvariable=self.frecventa)
        self.entry1.place(x=200, y=130)

        self.label2 = ttk.Label(self.tab1, text="Amplitudine : ",background='LightBlue1', width= 12, font=("arial",10,"bold"))
        self.label2.place(x=200, y=160)
        self.entry2=ttk.Entry(self.tab1,font=(30), textvariable=self.amplitudine)
        self.entry2.place(x=200, y=180)
        
        self.button1=tk.Button(self.tab1, text="Generare" ,bg='#BFEFFF', bd=0,width = 14, command=self.handle_sinus)
        self.button1.place(x=200, y=220)

        ###########################################################################################

        
        ########################        Ramp        ##################################
        
        self.label3 = ttk.Label(self.tab2, text="F.Es: ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label3.place(x=200, y=60)
        self.entry3=ttk.Entry(self.tab2,font=(30), textvariable=self.fesantionare)
        self.entry3.place(x=200, y=80)

        self.label4 = ttk.Label(self.tab2, text="Frecventa : ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label4.place(x=200, y=110)
        self.entry4=ttk.Entry(self.tab2,font=(30), textvariable=self.frecventa)
        self.entry4.place(x=200, y=130)

        self.label5 = ttk.Label(self.tab2, text="Amplitudine : ",background='LightBlue1', width= 12, font=("arial",10,"bold"))
        self.label5.place(x=200, y=160)
        self.entry5=ttk.Entry(self.tab2,font=(30), textvariable=self.amplitudine)
        self.entry5.place(x=200, y=180)
    
        
        self.button2=tk.Button(self.tab2, text="Generare",bg='#BFEFFF', bd=0,width = 14)
        self.button2['command'] = self.handle_ramp
        self.button2.place(x=200, y=220)


        ###########################################################################################


        ###################         Square       #######################################
        
        self.label6 = ttk.Label(self.tab3, text="F.Es: ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label6.place(x=200, y=60)
        self.entry6=ttk.Entry(self.tab3,font=(30), textvariable=self.fesantionare)
        self.entry6.place(x=200, y=80)

        self.label7 = ttk.Label(self.tab3, text="Frecventa : ",background='LightBlue1', width= 10, font=("arial",10,"bold"))
        self.label7.place(x=200, y=110)
        self.entry7=ttk.Entry(self.tab3,font=(30), textvariable=self.frecventa)
        self.entry7.place(x=200, y=130)

        self.label8 = ttk.Label(self.tab3, text="Amplitudine : ",background='LightBlue1', width= 12, font=("arial",10,"bold"))
        self.label8.place(x=200, y=160)
        self.entry8=ttk.Entry(self.tab3,font=(30), textvariable=self.amplitudine)
        self.entry8.place(x=200, y=180)
        
        self.button3=tk.Button(self.tab3, text="Generare",bg='#BFEFFF', bd=0,width = 14 ,command=self.start_square)
        self.button3.place(x=200, y=220)


        ###########################################################################################

        ###########################################################################################

        
        self.button4=tk.Button(self.header, text="Stop", bg='#BFEFFF', bd=0,width = 16)
        self.button4['command'] = self.stop_commands
        self.button4.place(x=85, y=290)
        
        
        self.button5=tk.Button(self.header, text="Save", bg='#BFEFFF', bd=0,width = 16)
        self.button5['command']= self.open_excel
        self.button5.place(x=250, y=290)
    def open_excel(self):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True
        excel.ScreenUpdating = True
        excel.DisplayAlerts = False
        excel.EnableEvents = True
        wa = excel.Workbooks.Open(self.fullpath)
        

    def save_excel(self):
        now = datetime.datetime.now()
        wb = win32.gencache.EnsureDispatch('Excel.Application')
        
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

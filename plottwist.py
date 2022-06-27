from tkinter import *
import numpy as np
from scipy import signal
import math
from drawnow import *
import matplotlib.pyplot as plt
import random
import time
import ADS1256
import DAC8532
import RPi.GPIO as GPIO

def sinus(fesantionare, fsemnal, amplitudine):
    fs = int(fesantionare) # frecventa esantionare
    A = amplitudine
    f0 = int(fsemnal) #freventa semnal
    tstep = 1/fs # interval timp esantion
    N = int(fs/f0)
    t = np.linspace(0, (N-1)*tstep, N)
    fstep = fs/N
    f=np.linspace(0, (N-1)*fstep, N)
    y = A*np.sin(2*np.pi*f0*t)
    return y

def ramp(fesantionare, fsemnal, amplitudine):
    fs = int(fesantionare) # frecventa esantionare
    A = amplitudine
    f0 = int(fsemnal) #freventa semnal
    tstep = 1/fs # interval timp esantion
    N = int(fs/f0)
    t = np.linspace(0, (N-1)*tstep, N)
    fstep = fs/N
    f=np.linspace(0, (N-1)*fstep, N)
    z = np.abs(A*signal.sawtooth(2*np.pi*f0*t))
    return z
    
def square(fesantionare, fsemnal, amplitudine):
    fs = int(fesantionare) # frecventa esantionare
    A = amplitudine
    f0 = int(fsemnal) #freventa semnal
    tstep = 1/fs # interval timp esantion
    N = int(fs/f0)
    t = np.linspace(0, (N-1)*tstep, N)
    fstep = fs/N
    f=np.linspace(0, (N-1)*fstep, N)
    x = A*signal.square(2*np.pi*f0*t)
    return x

semnal = square(1000,100,3)
vector =[]
count = 0

def makefig():
    plt.ylim(-5,5)
    plt.title("graphic")
    plt.grid(True)
    plt.ylabel("RandomNumber")
    plt.plot(vector, color='red', marker='o', markerfacecolor='blue', label='numere')
    plt.legend(loc='upper left')
    
while True:
    vector.append(semnal[count])    
    drawnow(makefig)
    #plt.pause(.0001)
    count=count+1
    print(count)
    if len(semnal) > 100:
        if count>100:
            vector.pop(0)
    if count>=len(semnal):
        vector.pop(0)
        count=0
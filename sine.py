from tkinter import *
import math
from drawnow import *
import matplotlib.pyplot as plt
import random
vector =[]
count = 0
def makefig():
    plt.ylim(0,30)
    plt.title("graphic")
    plt.grid(True)
    plt.ylabel("RandomNumber")
    plt.plot(vector, color='red', marker='o', markerfacecolor='blue', label='numere')
    plt.legend(loc='upper left')
while True:
    randomnumber = random.random()
    print(randomnumber)

    vector.append(randomnumber)
    drawnow(makefig)
    plt.pause(.0001)
    count=count+1
    if count>50:
        vector.pop(0)

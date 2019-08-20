# -*- coding: utf-8 -*-
"""
Created on Tue Oct 30 19:36:34 2018

@author: DELL
"""

from tkinter import * 
from tkinter import messagebox
import speech_recognition as sr #speech recognition
import pyttsx3 
import time
import openpyxl as op
import os

win=Tk()

win.geometry("1360x780")

def validatemail():
    if(len(entry4.get())!=0):
        is_valid = validate_email(entry4.get())
        if(is_valid==True):
            mailotp()
            #validatephone()
        else:
            engine = pyttsx3.init()
            engine.say("Invalid Email ID")#speak the text
            engine.runAndWait()
        
            messagebox.showerror("Error","Invalid Email ID")
    else:
        engine = pyttsx3.init()
        engine.say("Email cannot be blank")#speak the text
        engine.runAndWait()
        
        messagebox.showwarning("Warning","Email ID cannot be blank")

def submit():
    wb = op.load_workbook("C:/Users/Shani Singh/Desktop/Feedback.xlsx")
    ws = wb.active
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 30
    ws.append([entry2.get(),entry3.get(),entry4.get(),entry5.get()])
    wb.save("C:/Users/Shani Singh/Desktop/Feedback.xlsx")
    wb.close()

def ready1():
    try:
        engine = pyttsx3.init()
        engine.say("Hey Let's start filling the form")#speak the text
        engine.runAndWait()
        time.sleep(1)#wait    
        engine.say("So Your good name is")
        engine.runAndWait()
        time.sleep(0)#wait    
        
        r=sr.Recognizer()
        with sr.Microphone() as source:
            audio=r.listen(source,0,3)#listen the speaker
            name=r.recognize_google(audio)#converted to text
 
        time.sleep(0)#wait 
        
        entry2.delete(0, END)
        entry2.insert(10,name)

        engine = pyttsx3.init()
        engine.say("Your college name:")#speak the text
        engine.runAndWait()
    
        time.sleep(0)#wait    
        r=sr.Recognizer()#initializing the recogniser
        
        with sr.Microphone() as source:
            audio=r.listen(source,0,3)#listen the speaker
            college=r.recognize_google(audio)#converted to text
        
        entry3.delete(0, END)
        entry3.insert(10,college)
        
        engine = pyttsx3.init()
        engine.say("Branch")#speak the text
        engine.runAndWait()
        
        r=sr.Recognizer()#initializing the recogniser
        with sr.Microphone() as source:
            audio=r.listen(source,0,3)#listen the speaker
            branch=r.recognize_google(audio)#converted to text
        
        time.sleep(0)#wait
        
        entry4.delete(0, END)
        entry4.insert(10,branch)#insert the text on the respective entry field
    
        engine = pyttsx3.init()
        engine.say("Phone number")#speak the text
        engine.runAndWait()
        
        time.sleep(0)#wait    
        
        r=sr.Recognizer()#initializing the recogniser
        with sr.Microphone() as source:
            audio=r.listen(source,0,4)#listen the speaker
            phone=r.recognize_google(audio)#converted to text
        
        entry5.delete(0, END)
        entry5.insert(10,phone)#insert the text on the respective entry field
    
        engine = pyttsx3.init()
        engine.say("Email")#speak the text
        engine.runAndWait()
    
        time.sleep(0)#wait    
    
        r=sr.Recognizer()#initializing the recogniser
        with sr.Microphone() as source:
            audio=r.listen(source,0,3)#listen the speaker
            mail=r.recognize_google(audio)#converted to text
        
        mail=(mail.replace(" ",""))
        
        entry6.delete(0, END)
        entry6.insert(10,mail)#insert the text on the respective entry field
        
        engine = pyttsx3.init()
        engine.say("Do you want to edit anything?")#speak the text
        engine.runAndWait()
        
        t=messagebox.askquestion("Confirm","Do you want to edit anything?" )
        if(t=='yes'):
            button3=Button(win,text="EDIT",bg="pink")
            button3.place(x=230,y=100)
        
            button5=Button(win,text="EDIT",bg="pink")
            button5.place(x=230,y=180)
        else:
            edu_button.config(state="normal")
    except:
        engine = pyttsx3.init()
        engine.say("Unable to recognize your voice. Please try again")#speak the text
        engine.runAndWait()
        
        button3=Button(win,text="EDIT",bg="pink")
        button3.place(x=230,y=100)
        
        button5=Button(win,text="EDIT",bg="pink")
        button5.place(x=230,y=180)
       
def yes1():
    try:
        engine = pyttsx3.init()
        engine.say("Hey Let's start filling the form")#speak the text
        engine.runAndWait()
        time.sleep(1)#wait    
        engine.say("So Your good name is")
        engine.runAndWait()
        time.sleep(0)#wait    
        
        r=sr.Recognizer()
        with sr.Microphone() as source:
            audio=r.listen(source,0,3)#listen the speaker
            name=r.recognize_google(audio)#converted to text
 
        time.sleep(0)#wait 
        
        entry2.delete(0, END)
        entry2.insert(10,name)

        engine = pyttsx3.init()
        engine.say("Your college name:")#speak the text
        engine.runAndWait()
    
        time.sleep(0)#wait    
        r=sr.Recognizer()#initializing the recogniser
        
        with sr.Microphone() as source:
            audio=r.listen(source,0,3)#listen the speaker
            college=r.recognize_google(audio)#converted to text
        
        entry3.delete(0, END)
        entry3.insert(10,college)
    except:
        engine = pyttsx3.init()
        engine.say("Some Error occured please try again")#speak the text
        engine.runAndWait()
        
        messagebox.showwarning("Error","Some Error occured please try again")
        
def yes2():
    try:
        engine = pyttsx3.init()
        engine.say("Phone number")#speak the text
        engine.runAndWait()
        
        time.sleep(0)#wait    
        
        r=sr.Recognizer()#initializing the recogniser
        with sr.Microphone() as source:
            audio=r.listen(source,0,4)#listen the speaker
            phone=r.recognize_google(audio)#converted to text
        
        entry5.delete(0, END)
        entry5.insert(10,phone)#insert the text on the respective entry field
    
        engine = pyttsx3.init()
        engine.say("Email")#speak the text
        engine.runAndWait()
    
        time.sleep(0)#wait    
    
        r=sr.Recognizer()#initializing the recogniser
        with sr.Microphone() as source:
            audio=r.listen(source,0,3)#listen the speaker
            mail=r.recognize_google(audio)#converted to text
        
        mail=(mail.replace(" ",""))
        
        entry6.delete(0, END)
        entry6.insert(10,mail)#insert the text on the respective entry field
    except:
        engine = pyttsx3.init()
        engine.say("Some Error occured please try again")#speak the text
        engine.runAndWait()

        messagebox.showwarning("Error","Some Error occured please try again")
    
label=Label(win,text="REGISTRATION FORM",font="nones 14 bold")
label.place(x=500,y=10)

button1=Button(win,text="READY",bg="pink",command=ready1)
button1.place(x=720,y=10)

label2=Label(win,text="Name")
label2.place(x=10,y=60)
entry2=Entry(win)
entry2.place(x=80,y=60)

label3=Label(win,text="College")
label3.place(x=10,y=100)
entry3=Entry(win)
entry3.place(x=80,y=100)

label4=Label(win,text="Branch")
label4.place(x=10,y=140)
entry4=Entry(win)
entry4.place(x=80,y=140)

label5=Label(win,text="Phone No.")
label5.place(x=10,y=180)
entry5=Entry(win)
entry5.place(x=80,y=180)

label6=Label(win,text="Email")
label6.place(x=10,y=220)
entry6=Entry(win)
entry6.place(x=80,y=220)

button6=Button(win,text="SUBMIT",bg="pink",command=submit)
button6.place(x=230,y=260)

win.mainloop()

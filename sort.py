import pyttsx3 #texttovoicemodule
import speech_recognition as sr
import wikipedia
import webbrowser
import datetime
import os
import smtplib
import subprocess
import wolframalpha
import pyttsx3 
import tkinter 
import json 
import random 
import operator
import pyjokes  
import smtplib 
import ctypes 
import time 
import requests 
import shutil 
import twilio.rest
import client 
from bs4 import BeautifulSoup 
import win32com.client as wincl 
from urllib.request import urlopen 
import winshell  
import feedparser 
from twilio.rest import Client 
from clint.textui import progress 
from ecapture import ecapture as ec

engine = pyttsx3.init('sapi5') #MicrosoftspeakAPI
voices = engine.getProperty('voices') #print(voices[0].id)check voices available in code
engine.setProperty('voice', voices[0].id)


columns = shutil.get_terminal_size().columns
print ('********************************************************************************************************************************'.center(columns))
print ('********************************************************************************************************************************'.center(columns))
print ('INITIALIZING SORT'.center(columns))
print ('********************************************************************************************************************************'.center(columns))
print ('********************************************************************************************************************************'.center(columns))  
time.sleep(2)


def speak(audio): #speakfunctionof  Sort
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
            speak("good morning!")
    elif hour>=12 and hour<18:
            speak("good afternoon!")
    else:
            speak("good evening!")
        
    speak("Welcome")
    time.sleep(1)
    speak("My name is Sort ")
    speak("I will assist you here ")
    speak("What should i call you")
    uname = takeCommand()
    uname.upper() 
    speak("WELCOME"+uname.center(columns))
    print("WELCOME".center(columns)) 
    print(uname.upper().center(columns)) 
    speak("How can i Help you")
    time.sleep(1)

def takeCommand(): #it takes microphone voice as input and convert it into string output

        r = sr.Recognizer()
        with sr.Microphone() as source:
            print ('********************************************************************************************************************************'.center(columns))
            print("LISTENING".center(columns))
            print ('********************************************************************************************************************************'.center(columns))
            r.pause_threshold = 1 
            audio = r.listen(source)

        try:
            print ('********************************************************************************************************************************'.center(columns))
            print("RECOGNIZING".center(columns))
            print ('********************************************************************************************************************************'.center(columns))
            query = r.recognize_google(audio, language='en-in')
            
            
        
        except Exception as e:
            print(e) #don't print the error 
            print("CAN YOU SAY THAT AGAIN,PLEASE!".center(columns))
            return "None"
        return query


def sendEmail(to, content):
        server = smtplib.SMTP('smtp.gmail.com',587)
        server.ehlo()
        server.starttls()
        server.login('SenderEmail','EmailPassword')
        server.sendmail('SenderEmail',to,content)
        server.close()


if __name__ == "__main__":
    wishMe() 
    while True:
        query = takeCommand().lower() 


        if  'wikipedia' in query:
                speak('Searching on Wikipedia')
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences=2)
                speak("According to wikipedia")
                print(results.center(columns))
                speak(results)

        elif 'open youtube' in query:
                webbrowser.open("https://www.youtube.com")
                
        elif 'open google' in query:
                webbrowser.open("https://www.google.com")

        elif 'open facebook' in query:
                 webbrowser.open("https://www.facebook.com")

        elif 'open twitter' in query:
                webbrowser.open("https://www.twitter.com")

        elif 'open Linkedin' in query:
                webbrowser.open("https://www.linkedin.com")
                
        elif 'open amazon' in query:
                webbrowser.open("https://www.amazon.com")

        elif 'open flipkart' in query:
                webbrowser.open("https://www.flipkart.com")

        elif 'open whatsapp' in query:
                webbrowser.open("web.whatsapp.com")

        elif 'cricket' in query:
                webbrowser.open("https://www.cricbuzz.com")

        elif 'cricket score' in query:
                webbrowser.open("https://www.cricbuzz.com")

        elif 'cricket scores' in query:
                webbrowser.open("https://www.cricbuzz.com")

        elif 'open instagram' in query:
                webbrowser.open("https://www.instagram.com")

        elif 'open music' in query:
                webbrowser.open("https://music.youtube.com/watch?v=wdEl9h4dAKo&list=RDAMVMwdEl9h4dAKo")

        elif 'play music' in query:
                webbrowser.open("https://music.youtube.com/watch?v=wdEl9h4dAKo&list=RDAMVMwdEl9h4dAKo")

        elif 'play song' in query:
                webbrowser.open("https://music.youtube.com/watch?v=wdEl9h4dAKo&list=RDAMVMwdEl9h4dAKo")
                
        elif 'open song' in query:
                webbrowser.open("https://music.youtube.com/watch?v=wdEl9h4dAKo&list=RDAMVMwdEl9h4dAKo")
                
        elif 'sing a song' in query:
                webbrowser.open("https://music.youtube.com/watch?v=wdEl9h4dAKo&list=RDAMVMwdEl9h4dAKo")

        elif 'the time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak("The Time is")
                speak(strTime)
                print(strTime.center(columns))

        elif 'how are you' in query: 
                speak("I am fine, Thank you") 
                speak("How are you") 

        elif "what's your name" in query or "What is your name" in query: 
                speak("My friends call me") 
                speak("SORT") 
                print("My friends call me".center(columns)) 
                print("SORT".center(columns))

        elif 'exit' in query: 
                speak("Thanks for giving me your time") 
                exit() 

        elif "who made you" in query or "who created you" in query:  
                speak("I have been created by DOT SYSTEMS.center(columns)") 

        elif 'joke' in query: 
                speak(pyjokes.get_joke())
                print(pyjokes.get_joke().center(columns))    

        elif "who i am" in query: 
                speak("If you talk then definately your human.") 

        elif "why you came to world" in query: 
                speak("Thanks to DOT SYSTEMS. further It's a secret") 

        elif 'is love' in query: 
                speak("It is 7th sense that destroy all other senses") 

        elif "who are you" in query: 
                speak("I am your virtual assistant created by DOT SYSTEMS") 

        elif 'reason for you' in query: 
                speak("I was created as a Minor project by DOT SYSTEMS ")  

        elif 'lock window' in query: 
                speak("locking the device") 
                ctypes.windll.user32.LockWorkStation() 

        elif 'shutdown system' in query: 
                speak("Hold On a Sec ! Your system is on its way to shut down") 
                subprocess.call('shutdown / p /f') 

        elif 'empty recycle bin' in query: 
                winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True) 
                speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query: 
                speak("for how much time you want to stop SORT from listening commands") 
                a = int(takeCommand()) 
                time.sleep(a) 
                print(a) 

        elif "camera" in query or "take a photo" in query: 
                ec.capture(0, "SORT Camera ", "img.jpg") 
        
        elif "restart" in query: 
                subprocess.call(["shutdown", "/r"]) 
                
        elif "hibernate" in query or "sleep" in query: 
                speak("Hibernating") 
                subprocess.call("shutdown / h") 
        
        elif "log off" in query or "sign out" in query: 
                speak("Make sure all the application are closed before sign-out") 
                time.sleep(5) 
                subprocess.call(["shutdown", "/l"])

          

        elif "sort" in query: 
                wishMe() 
                speak("sort in your service ") 
                speak(uname) 

        elif "will you be my gf" in query or "will you be my bf" in query:    
                speak("I'm not sure about, may be you should give me some time") 

        elif "how are you" in query: 
                speak("I'm fine, glad you me that") 
                
        elif "i love you" in query: 
                speak("It's hard to understand") 

        elif 'change background' in query: 
                ctypes.windll.user32.SystemParametersInfoW(20,0,"Location of wallpaper",0)

        elif 'open virtual studio' in query:
                path = "D:Microsoft VS Code\Code.exe"
                os.startfile(path)

        
        elif 'send email to mam' in query:              #1
                try:
                        speak("What should I mail her?")
                        content= takeCommand()
                        to = "RecieverEmail"
                        sendEmail(to,content)
                        speak("Email has been successfully send to her!")
                        print("Email has been successfully send to her!".center(columns))

                except Exception as e:
                                print(e)
                                speak("Sorry Email is not send please check and let me know")

        elif 'send email to mam' in query:             #2
                try:
                        speak("What should I mail her?")
                        content= takeCommand()
                        to = "RecieverEmail2"
                        sendEmail(to,content)
                        speak("Email has been successfully send to her!")
                        print("Email has been successfully send to her!".center(columns))


                except Exception as e:
                        print(e)
                        speak("Sorry Email is not send please check and let me know")

        elif 'send email to mam' in query:         #1.1
                try:
                        speak("What should I mail her?")
                        content= takeCommand()
                        to = "RecieverEmail2"
                        sendEmail(to,content)
                        speak("Email has been successfully send to her!")
                        print("Email has been successfully send to her!".center(columns))


                except Exception as e:
                        print(e)
                        speak("Sorry Email is not send please check and let me know")

                
        elif 'send email to mam' in query:         #2.1
                try:
                        speak("What should I mail her?")
                        content= takeCommand()
                        to = "RecieverEmail"
                        sendEmail(to,content)
                        speak("Email has been successfully send to her!")
                        print("Email has been successfully send to her!".center(columns))
                
                
                except Exception as e:
                        print(e)
                        speak("Sorry Email is not send please check and let me know")


        elif 'send a mail' in query:            #3
                try: 
                        speak("What should I say?") 
                        content = takeCommand() 
                        speak("whome should i send") 
                        to = input()  
                        sendEmail(to, content) 
                        speak("Email has been successfully send!")
                        print("Email has been successfully send!".center(columns))


                except Exception as e: 
                        print(e) 
                        speak("Sorry Email is not send please check and let me know") 



                        
                
                
                






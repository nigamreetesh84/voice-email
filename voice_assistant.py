import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
#from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import re

from gmail import Gmail

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning  !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon  !")

    else:
        speak("Good Evening  !")

    # assname = ("Pihu")
    speak("I am your Assistant")
    # speak(assname)


def usrname():
    # speak("What should i call you ")
    # uname = takeCommand()
    # speak("Welcome to voice assistant")
    # speak(uname)
    # columns = shutil.get_terminal_size().columns

    # print("#####################".center(columns))
    # print("Welcome Mr.", uname.center(columns))
    # print("#####################".center(columns))

    speak("How can i Help you, ")


def takeCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:
        # r.adjust_for_ambient_noise(source, duration=5)
        # r.dynamic_energy_threshold = True
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print("...")
        # print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query


if __name__ == '__main__':
    def clear(): return os.system('cls')

    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    usrname()
    repeat = False
    pause = False

    while True:

        # query = takeCommand().lower()
        r = sr.Recognizer()
        mic = sr.Microphone(device_index=0)
        with mic as source:
            # r.adjust_for_ambient_noise(source, duration=1)
            # r.dynamic_energy_threshold = True
            r.energy_threshold = 400
            r.pause_threshold = 2  # any error comes then just commment this one SIDDHU
            # print(recognizer.energy_threshold,"Threshold")
            r.dynamic_energy_threshold = False
            print("Listening your instruction...")
            # audio=recognizer.listen(source,timeout=5) # removing timeout as this is not suitable for me now, Siddhu
            audio = r.listen(source)
            print("*=*"*5)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language='en-in')
            #speak(query)
            print(f"User said: {query}\n")
        except Exception as e:
            print(e)
            print("Unable to Recognize your voice.")
            continue

        if re.search("pause|wait", query):
            #speak("I am sleeping")
            time.sleep(5)
            print("I am in sleep")
            pause = True
            continue

        if "start voice" in query:
            print("going to start")
            pause = False

        elif not pause:
            print("Start recording")
            try:
                if 'wikipedia' in query:
                    speak('Searching Wikipedia...')
                    query = query.replace("wikipedia", "")
                    results = wikipedia.summary(query, sentences=3)
                    speak("According to Wikipedia")
                    print(results)
                    speak(results)

                if "lock" in query:
                    speak("Going to lock system")
                    ctypes.windll.user32.LockWorkStation()
                    time.sleep(2)
                    speak("Reetesh, you have meeting in 15 mins")
                    speak("Subject is load balancer pre production deployment")
                    speak("Any action item")
                    content = takeCommand()
                    content = content.split("reply")[-1]
                    mail_obj = Gmail()
                    sender= "nigamreetesh84@gmail.com",
                    to="nigamreetesh84@yahoo.com",
                    subject=" Load balancer pre production deployment",
                    msg_html=content,
                    msg_plain=content,
                    rsp = mail_obj.send_message(sender= "nigamreetesh84@gmail.com",
                        to="nigamreetesh84@yahoo.com",
                        subject="Load balancer pre production deployment",
                        msg_html=content
                    )
                    speak("successfully email sent")
                
            except Exception as e:
                print(e)

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

                elif 'deploy' in query:
                    speak("welcome to saturn cloud, which project do you want to \
                        deploy?")
                    project = takeCommand()
                    speak(
                        "Going to get deployment detials for project {0}".format(project))
                    speak("Any extra action if build successfully !")
                    success = takeCommand()
                    speak("Any extra action if build fails !")
                    fail = takeCommand()
                    speak("Will excute job, take the action as per you mentioned\
                        if fail will {0} and success will {1} ".format(
                        fail, success))
                    rsp = send_email.send(file_name='etl-image.png')
                    speak("successfully email sent")

                elif 'pipeline' in query:
                    speak("There are 4 pipeline jobs configured, please provide more \
                        details")
                    job_id = takeCommand()
                    with open(os.path.join(os.getcwd(), "etl", "readme.md")) as fd:
                        # for data in fd.readlines():
                        speak(fd.read())
                    speak("Any extra action if job runs successfully !")
                    success = takeCommand()
                    speak("Any extra action if job fails !")
                    fail = takeCommand()
                    speak("Will excute job, take the action as per you mentioned if fail will {0} and success will {1} ".format(
                        fail, success))
                    speak("Any more jobs you want to run")

                elif 'servicenow' in query:
                    speak("Please provide following details")
                    speak("what should be the urgency value, Low, Medium, High")
                    urgency = takeCommand().lower()
                    speak("Provide short decription")
                    short_description = takeCommand().lower()
                    if not urgency:
                        urgency = "medium"
                    res = incident.post_incident_request(
                        urgency=urgency, short_description=short_description)
                    speak(res)

                elif "architecture" in query or "draw" in query:
                    query = query.split("architecture diagram")
                    arch = query[1].split("and")
                    speak("Do you want to provide the details or shall I draw for demo")
                    data = takeCommand()
                    speak("Going to draw architecture diagram {0}".format(
                        " ".join(arch[1:])))
                    draw_arch.generate_diagram()

                elif 'open youtube' in query:
                    speak("Here you go to Youtube\n")
                    webbrowser.open("youtube.com")

                # elif 'open google' in query:
                #     speak("Here you go to Google\n")

                #     webbrowser.open("google.com")

                elif 'open stackoverflow' in query:
                    speak("Here you go to Stack Over flow.Happy coding")
                    webbrowser.open("stackoverflow.com")

                elif 'play music' in query or "play song" in query:
                    speak("Here you go with music")
                    # music_dir = "G:\\Song"
                    music_dir = "C:\\songs"
                    songs = os.listdir(music_dir)
                    print(songs)
                    speak("Please select your choice")
                    for i in songs:
                        speak(i)
                    keyword = takeCommand().lower()
                    if keyword in songs:
                        speak("Enjoy this {} song ".format(keyword))
                        song = os.listdir(os.path.join(music_dir, keyword))[0]
                        webbrowser.open(os.path.join(music_dir, keyword, song))
                        time.sleep(2)
                    #random = os.startfile(os.path.join(music_dir, songs[1]))

                elif 'the time' in query:
                    strTime = datetime.datetime.now().strftime("% H:% M:% S")
                    speak(f", the time is {strTime}")

                elif 'send a mail' in query:
                    try:
                        speak("What should I say?")
                        content = takeCommand()
                        speak("whome should i send")
                        to = input()
                        sendEmail(to, content)
                        speak("Email has been sent !")
                    except Exception as e:
                        print(e)
                        speak("I am not able to send this email")

                elif 'exit' in query or 'stop' in query or 'bye' in query:
                    speak("Thanks for giving me your time")
                    exit()

                elif 'joke' in query:
                    speak(pyjokes.get_joke())

                elif "calculate" in query:

                    app_id = "KPGL8P-YAKX2J5RAK"
                    client = wolframalpha.Client(app_id)
                    indx = query.lower().split().index('calculate')
                    query = query.split()[indx + 1:]
                    res = client.query(' '.join(query))
                    answer = next(res.results).text
                    print("The answer is " + answer)
                    speak("The answer is " + answer)

                elif "google search" in query:
                    speak("What should I search in google?")
                    keyword = takeCommand()
                    if keyword:
                        url = "https://google.com/search?q=" + keyword

                        # webbrowser module to work with the webbrowser
                        speak("Here are the search results for " + keyword)
                        webbrowser.open(url)

                elif 'news' in query:
                    try:
                        jsonObj = urlopen(
                            '''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=5528632365d2438c9fe1efcd2ef64419''')
                        data = json.load(jsonObj)
                        i = 1

                        speak('here are some top news from the times of india')
                        print(
                            '''=============== TIMES OF INDIA ============''' + '\n')

                        for item in data['articles']:

                            print(str(i) + '. ' + item['title'] + '\n')
                            print(item['description'] + '\n')
                            speak(str(i) + '. ' + item['title'] + '\n')
                            i += 1
                            if i == 4:
                                speak("Do you want to listen more news")
                                keyword = takeCommand()
                                if keyword in ("yes", "continue", "keep"):
                                    pass
                                else:
                                    break
                    except Exception as e:

                        print(str(e))

                elif query in ('show note', 'read note', 'show notes', 'read notes'):
                    speak("Showing Notes")
                    with open("test.txt", "r") as fd:
                        data = fd.read()
                        print(data)
                        speak(str(data))

                elif "note" in query:
                    note_status = True
                    with open('test.txt', 'w') as fd:
                        speak("Should i include today's date ?")
                        snfm = takeCommand()
                        strTime = "{:%B %d, %Y}".format(
                            datetime.datetime.now())
                        if 'yes' in snfm or 'sure' in snfm:
                            # datetime.datetime.now().strftime("%m/%d/%Y, %H:%M:%S")
                            strTime = "{:%B %d, %Y}".format(
                                datetime.datetime.now())
                            fd.write(strTime)
                        # while note_status:
                        note = takeCommand()
                        data = note.split("i am done")
                        body = "Notes"
                        subject = "notes " + strTime
                        # if re.search(r'i am done|complete|done|finish', note):
                        #     print("Cloing the file")
                        #     note_status = False
                        #     fd.close()
                        # elif note:
                        fd.write(str(data[0]))
                        print("Notes written in file")
                    speak("Shall I send this to your gmail account?")
                    cmd = takeCommand()
                    if "yes" in cmd or "please" in cmd:
                        print("Going to send email")
                        rsp = send_email.send(file_name='test.txt')
                        speak("successfully email sent")

                elif "weather" in query:
                    # Google Open weather website
                    # to get API of Open weather
                    api_key = "5325b12a3a49d95676d8bb5bddf64565"
                    base_url = "http://api.openweathermap.org/data/2.5/weather?"
                    speak(" City name ")
                    print("City name : ")
                    city_name = takeCommand()
                    complete_url = base_url + "appid=" + api_key + "&q=" + city_name
                    print(complete_url)
                    response = requests.get(complete_url)
                    x = response.json()

                    if x["cod"] != "404":
                        y = x["main"]
                        current_temperature = y["temp"]
                        current_pressure = y["pressure"]
                        current_humidiy = y["humidity"]
                        z = x["weather"]
                        weather_description = z[0]["description"]
                        print(" Temperature (in kelvin unit) = " + str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(
                            current_pressure) + "\n humidity (in percentage) = " + str(current_humidiy) + "\n description = " + str(weather_description))

                    else:
                        speak(" City Not Found ")

                elif "send message " in query:
                    # You need to create an account on Twilio to use this service
                    account_sid = 'Account Sid key'
                    auth_token = 'Auth token'
                    client = Client(account_sid, auth_token)

                    message = client.messages \
                        .create(
                            body=takeCommand(),
                            from_='Sender No',
                            to='Receiver No'
                        )

                    print(message.sid)

                elif "wikipedia" in query:
                    webbrowser.open("wikipedia.com")

                elif "what is" in query or "who is" in query:

                    # Use the same API key
                    # that we have generated earlier
                    client = wolframalpha.Client("KPGL8P-YAKX2J5RAK")
                    res = client.query(query)

                    try:
                        print(next(res.results).text)
                        speak(next(res.results).text)
                    except StopIteration:
                        print("No results")

                # elif "" in query:
                    # Command go here
                    # For adding more commands

            except Exception as e:
                print(e)

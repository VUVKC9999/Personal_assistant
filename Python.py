import smtplib
import subprocess
from winreg import QueryValue
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
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
# can change the voice by changung the number in abouve line

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def WishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour < 12:
        speak("Good Morning sir !")

    elif hour>=12 and hour<18:
        speak("Good Afternoon sir !")

    else:
        speak("Good Evening sir ! ")

    assname = ("Jarvis")
    speak("I am Your Personal Assistant")
    speak(assname)


def username():
    speak("What Should I call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print("########################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#########################".center(columns))

    speak("How can I Help you")

def takeCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening....")
        r.pause_threshold = 1
        audio = r.listen(source)
    
    try:
        print("Recognizing....")
        query = r.recognize_google(audio, language = 'en-in')
        print(f"user said :{query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    #Enable low security in gmail
    server.login('Your Email ID:', 'Your Email Password')
    server.sendmail('Your Email ID', to, content)
    server.close()


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    #this function will clean any command before execution of this python file
    clear()
    WishMe()
    username()

    while True:

        query = takeCommand().lower()
        #All the command said by user will be stored here in query

        if 'wikipedia' in query:
            speak('Searching Wikipedia....')
            query = query.replace("Wikipedia", "  ")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'play' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over Flow")
            webbrowser.open("stackoverflow.com")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            speak(f"Boss, the time is {strTime}")

        elif 'open opera' in query:
            codePath = r"C:\Users\karth\OneDrive\Desktop\Opera Browser.lnk"
            os.startfile(codePath)

        elif 'koti' in query:
            speak("He is a stupid guy, I don't want to talk about him")

        elif 'send a mail' in query:
            try:
                speak("What should I send")
                content = takeCommand()
                speak("Whome Should I send")
                to = input()
                sendEmail(to, content)
                speak("Email has sent !")
            except Exception as e:
                print(e)
                speak("I an unable to send Email")

        elif 'how are you' in query:
            speak("Iam fine, Thank you")
            speak("How are you, Boss")

        elif 'I am fine' in query or " I am good" in query:
            speak("It's good to know that you are fine")

        elif "Change my name to" in query:
            query = query.replace("Change my name to", " ")
            assname = query

        elif "change your name " in query:
            speak("What would you like to cakk me, Boss ? ")
            assname = takeCommand()
            speak("Thanks for naming me ")

        elif "What is your name" in query or "What's your name " in query:
            speak("Generally everyone uses to call me as")
            speak(assname)
            print("Generally everone uses to call me as", assname)

        elif "stop" in query or "exit" in query:
            speak("Thanks for giving me your time")
            exit()

        elif "Who made you" in query or "Who created you" in query:
            speak(" I am created by Karthik and Mourya")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:
            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            speak("The answer is " + answer)
            print("The answer is " + answer)

        elif 'search' in query or 'play' in query:
            query = query.replace("Search", " ")
            query = query.replace("play", " ")
            webbrowser.open(query)

        elif "Who am I " in query:
            speak("from the research I did, i came to find something weird, but as you are able to speak, I conclude that you are Human")

        elif "Why you came to world" in query:
            speak("For this I should first thank Karthik and Mourya, Further It's a top secret")

        elif 'power point presentation' in query:
            speak("Opening Power Point Presentation")
            power = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\LibreOffice 7.3"
            os.startfile(power)

        elif 'what is love' in query:
            speak("It Is 7th sense which destroys all other senses")

        elif 'who are you' in query:
            speak("I am an virtual Assistant, To serve you")

        elif 'reason for you' in query:
            speak("The reason I was created is to serve people and  help them in daily works")

        elif 'news' in query:
            try:
                jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1

                speak("Here are some top news from teh times of india")
                print('''=============== TIMES OF INDIA ============'''+ '\n')

                for item in data['articles']:
                    print(str(i) + '.' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '.' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                print(str(e))

        elif 'lock windows' in query:
            speak("Locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'Shutdown System' in query:
            speak("Wait a while Boss, Shutting down your System")
            subprocess.call('shutdown / p / f')

        elif 'empty my recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("clearing recycle bin")

        elif "don't listen" in query or "stop listening" in query:
            speak("from how much time you want me to stop listening your commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("Where is", " ")
            location = query
            speak("User asked to locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif 'open camera' in query or "take a photo" in query:
            ec.capture(0,"Jarvis Camera ", "imp.jpg")

        elif 'restart' in query:
            subprocess.call(["shutdown", "/r"])

        elif 'sleep' in query:
            speak("Changing the system to Sleep mode")
            subprocess.call("shutdown / h")

        elif "Write a notes" in query:
            speak("What should I write, Boss")
            note = takeCommand()
            file = open('Jarvis.txt', 'w')
            speak("Boss, Should I include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)    

        elif 'show notes' in query:
            speak("Showing notes")
            file = open("Jarvis.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif 'update assistant' in query:
            speak("After downloading replace me with the newely downloaded file")
            url = '# url after uploading file'
            r = requests.get(url, stream = True)

            with open("Voice.py", "wb") as pypdf:
                total_length = int(r.headers.get('content-length'))

                for ch in progress.bar(r.iter_content(chunk_size = 2391975), expected_size = (total_length / 1024) + 1):
                    if ch:
                        pypdf.write(ch)

        elif "Jarvis" in query:
            WishMe()
            speak("Jarvis in your service Boss")
            speak(assname)

        elif "weather" in query:
             
            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()
             
            if x["code"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))

            else:
                speak("City Not Found")

        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Boss")
            speak(assname)

        elif "will you be my gf " in query or "will you be my bf" in query:
            speak("I am not sure, I need some time to say")

        elif "How are you" in query:
            speak("I am fine, What about you")
            ans = takeCommand()
            if 'fine' in ans or 'I am good' in ans:
                speak("Glad to know that you are doing good")
            else:
                speak("Sorry")

        elif "I love you" in query:
            speak("I Love you too, but I am not in love with you")

 #       elif "what is" in query or "Who is " in query:
  #          client = wolframalpha.Client("API_ID")
   #         res = client.query(query)

 #           try:
#                print(next(res.results).text)
 #               speak(next(res.results).text)

  #          except StopIteration:
   #             print("No results")
    #            speak("Sorry to say this, I found no results related to that")

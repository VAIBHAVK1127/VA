import cgitb
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
import pywhatkit
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
import pyautogui
import keyboard as k
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    assname = ("Nezuko")
    speak("I am your Assistant")
    speak(assname)


def username():
    speak("How should I address you")
    uname = takeCommand()
    speak("Welcome")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print("Welcome ", uname)

    speak("How can I Help you, Sir")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return query


def sendEmail(to, subject, content, filename):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    server.login('vvk001127@gmail.com', 'ysczyqfhmhroasax')

    msg = MIMEMultipart()
    msg['From'] = 'vvk001127@gmail.com'
    msg['To'] = to
    msg['Subject'] = subject

    body = MIMEText(content, 'plain')
    msg.attach(body)
    attachment = open(filename, 'rb')
    base_attachment = MIMEBase('application', 'octet-stream')
    base_attachment.set_payload((attachment).read())
    encoders.encode_base64(base_attachment)
    base_attachment.add_header("Content-Disposition", f"attachment; filename= {filename}")
    msg.attach(base_attachment)
    text = msg.as_string()
    server.sendmail('vvk001127@gmail.com', to, text)
    server.quit()
    speak("Email has been sent !")


if __name__ == '__main__':
    clear = lambda: os.system('cls')

    clear()
    wishMe()
    username()

    while True:

        query = takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif ' message to mother' in query or 'message to mom' in query:
            speak("whats the message\n")
            query=takeCommand()
            pywhatkit.sendwhatmsg_instantly('+919594474009',query)
            pyautogui.click(1331, 700)
            time.sleep(2)
            k.press_and_release('enter')

        elif ' photo to mother' in query or 'pic to mom' in query:
           imagepath = "C:/Users/DELL/Downloads/AA.PNG"
           caption= "photo"
           pywhatkit.sendwhats_image('+919594474009',imagepath,caption,20)


        elif 'youtube' in query:
            query=query.replace('play',"")
            speak("playing\n" + query)
            pywhatkit.playonyt(query)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")



        elif ' my music' in query or " my song" in query:
            speak("Here you go with music")
            music_dir = "C:\\Users\\DELL\\Music\\music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[0]))


        elif 'mail to Vaibhav' in query or 'email to vaibhav' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "vaibhavkhrt7@gmail.com"
                subject = "Email from Assistant"
                filename = "C:/Users/DELL/Downloads/1.pdf"
                sendEmail(to, subject, content, filename)
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'send  mail' in query or 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                to = takeCommand() + "@gmail.com"
                subject = "Email from Assistant"
                filename = "C:/Users/DELL/Downloads/1.pdf"
                sendEmail(to, subject, content, filename)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Mister Vaibahv.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:

            app_id = "Wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'search' in query :
            query = query.replace("search", "")
            webbrowser.open(query)

        elif "who i am" in query:
            speak("If you talk then definitely your human.")

        elif "why you came to world" in query:
            speak("All Thanks to Mister Vaibhav. Rest of the part is confidential")

        elif 'power point' in query:
            speak("opening Power Point")
            power = r"C:\Users\DELL\Documents\msc part2\blockchain"
            os.startfile(power)

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your virtual assistant created by Mister Vaibhav")

        elif 'reason for you' in query:
            speak(
                "I was created as a Minor project by Mister Vaibhav when he completed his python course from Patkar College.")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 0, "C:\\Users\\DELL\\Downloads", 0)
            speak("Background changed successfully")


        elif 'news' in query:

            try:
                jsonObj = urlopen(
                    "https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\")
                data = json.load(jsonObj)
                i = 1

                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============''' + '\n')

                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:

                print(str(e))


        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop Nezuko from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query or  "shut down" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Nezuko Camera ", "img.jpg")

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('Note.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("Note.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif "update assistant" in query:
            speak("After downloading file please replace this file with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream=True)

            with open("Voice.py", "wb") as Pypdf:

                total_length = int(r.headers.get('content-length'))

                for ch in progress.bar(r.iter_content(chunk_size=2391975),
                                       expected_size=(total_length / 1024) + 1):
                    if ch:
                        Pypdf.write(ch)

        elif "Nezuko" in query:

            wishMe()
            speak("Nezuko at your service ")
            speak(assname)

        elif "weather" in query:

            api_key = "Api key"
            base_url = "https://chrome.google.com/webstore/detail/weather/iolcbmjhmpdheggkocibajddahbeiglb?hl=en"
            speak("City Name")
            print("City Name:")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " + str(
                    current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                    current_pressure) + "\n humidity (in percentage) = " + str(
                    current_humidiy) + "\n description = " + str(weather_description))

            else:
                speak(" City Not Found ")



        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")

        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(assname)


        elif "how are you" in query:
            speak("I'm fine, glad you me that")



        elif "what is" in query or "who is" in query:

            client = wolframalpha.Client("API_ID")
            res = client.query(query)

            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")
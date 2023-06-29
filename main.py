import subprocess
import wolframalpha
import tkinter
import json
import random
import operator
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
import pyaudio
import struct
import azure.cognitiveservices.speech as speechsdk
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

def speak(text: str):
    speech_config = speechsdk.SpeechConfig(subscription=os.environ.get('SPEECH_KEY'),
                                           region=os.environ.get('SPEECH_REGION'))

    audio_config = speechsdk.audio.AudioOutputConfig(use_default_speaker=True)

    speech_config.speech_synthesis_voice_name = 'en-IN-NeerjaNeural'

    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)

    speech_synthesis_result = speech_synthesizer.speak_text_async(text).get()

    if speech_synthesis_result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        print("Speech synthesized for text [{}]".format(text))
    elif speech_synthesis_result.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = speech_synthesis_result.cancellation_details
        print("Speech synthesis cancelled: {}".format(cancellation_details.reason))
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            if cancellation_details.error_details:
                print("Error details: {}".format(cancellation_details.error_details))
                print("Did you set the speech resource key and region values?")

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")
    else:
        speak("Good Evening Sir !")

    ainame = ("Baymax 1 point o")
    speak("I am your personal assistant")
    speak(ainame)

def username():
    speak("What should i call you sir?")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print("#####################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#####################".center(columns))

    speak("How can i help you, sir")

def takeCommand():
    speech_config = speechsdk.SpeechConfig(subscription=os.environ.get('SPEECH_KEY'),
                                           region=os.environ.get('SPEECH_REGION'))
    speech_config.speech_recognition_language = "en-IN"  # en-US

    audio_config = speechsdk.audio.AudioConfig(use_default_microphone=True)
    speech_recognizer = speechsdk.SpeechRecognizer(speech_config=speech_config, audio_config=audio_config)

    print("Speak into your microphone.")
    speech_text = speech_recognizer.recognize_once_async().get()
    transcript = format(speech_text.text)


    if speech_text.reason == speechsdk.ResultReason.RecognizedSpeech:
        print("Recognized: ", transcript)
    elif speech_text.reason == speechsdk.ResultReason.NoMatch:
        print("No speech could be recognized: {}".format(speech_text.no_match_details))
    elif speech_text.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = speech_text.cancellation_details
        print("Speech Recognition canceled: {}".format(cancellation_details.reason))
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print("Error details: {}".format(cancellation_details.error_details))
            print("Did you set the speech resource key and region values?")

    return transcript

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    server.login('your_email', 'your_password')
    server.sendmail('your_email', to, content)
    server.close()

def main():
    clear = lambda : os.system('cls')

    clear()
    wishMe()
    username()

    while True:

        query = takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query,sentences= 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to YouTube \n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("here you go to google \n")
            webbrowser.open("google.com")
        elif 'open stackoverflow' in query:
            speak("Here you go to stackoverflow")
            webbrowser.open("stackoverflow.com")
        elif 'play music' or 'play song' in query:
            speak("Enjoy your music sir")

            music_dir = "C:\\Users\\username\\Music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[1]))

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            speak(f"Sir, the time is {strTime}")

        elif 'open_quora' in query:
            codePath = "enter your exe file path"
            os.startfile(codePath)

        elif 'email to a person' in query:
            try:
                speak("okay, what should i say?")
                content = takeCommand()
                to = "receiver's email address"
                sendEmail(to, content)
                speak("Email has been sent sir")

            except Exception as e:
                print(e)
                speak("I am not able to send this email, sir")

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("Whom should i send this to?")
                to = input("Enter the receiver's email address: ")
                sendEmail(to, content)
                speak("Email has been sent sir")

            except Exception as e:
                print(e)
                speak("i'm not able to send this email, sir")

        elif 'how are you' in query:
            speak("I am doing good")
            speak("How about you, sir")

        elif 'fine' or 'good' in query:
            speak("I'm glad")

        elif 'change my name' in query:
            speak("What would you like to call me sir?")
            ainame = takeCommand()
            speak("thanks for naming me, sir")

        elif "what's your name" or "what is your name" in query:
            speak("My friends call me")
            speak(ainame)
            print("My friends call me", ainame)

        elif 'exit' in query:
            speak("thanks for giving me your time sir")
            exit()

        elif "who made you" or "who created you" in query:
            speak("I had been created by you sir, Mister Shriharan")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif 'calculate' in query:

            app_id = "APP_ID" #wolframaplha app id
            client = wolframalpha.Client(app_id)
            index = query.lower().split().index('calculate')
            query = query.split()[index + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'search' in query or 'play' in query:

            query = query.replace("search", "")
            query = query.replace("google", "")
            webbrowser.open(query)

        elif "who i am" in query:
            speak("If you talk then definitely you are a human")

        elif "why you came to world" in query:
            speak("Thanks to Shriharan. Furtherm it's a secret")

        elif "power point presentation" in query:
            speak("Opening power point presentation")
            power = r"C:\\Users\\username\\Desktop\\Baymax\\presentation.pptx"
            os.startfile(power)

        elif "is love" in query:
            speak("It's the 7th sense that destroys all other senses")

        elif "who are you" in query:
            speak("i am your virtual assistant created by you sir")

        elif "reason for you" in query:
            speak("I was created as a minor project by Shriharan")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "Location of wallpaper",
                                                       0)
            speak("Background changed successfully")

        elif 'news' in query:

            try:
                jsonObj = urlopen(
                    '''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\newsapi api key\\''')
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

        elif "don't listen" or "stop listening" in query:
            speak("for how much time you want to stop baymax from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Baymax Camera ", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('baymax.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("baymax.txt", "r")
            print(file.read())
            speak(file.read(6))

        elif "baymax" in query:

            wishMe()
            speak("Baymax 1 point o in your service Mister")
            speak(ainame)

        elif "weather" in query:
            
            api_key = "api key" #openweathermap api key
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
                print(" Temperature (in kelvin unit) = " + str(
                    current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                    current_pressure) + "\n humidity (in percentage) = " + str(
                    current_humidiy) + "\n description = " + str(weather_description))

            else:
                speak(" City Not Found ")

        elif "send message" in query:

            #create twilio account
            account_sid = 'account_sid'
            auth_token = 'auth_token'
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

        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(ainame)

        elif "will you be my gf" in query or "will you be my bf" in query:
            speak("I'm not sure about, may be you should give me some time")

        elif "how are you" in query:
            speak("I'm fine, glad you me that")

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what is" in query or "who is" in query:

            client = wolframalpha.Client("api key") #wolframalpha api key
            res = client.query(query)

            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")


        # elif "" in query:
        # Command go here
        # For adding more commands

if __name__ == '__main__':
    main()

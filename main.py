'''
Project: Jarvis AI
Name: Aditya Mangal

'''

import speech_recognition as sr  # pip install speechRecognition
from win32com.client import Dispatch # pip install pywin32
import datetime
import webbrowser
from termcolor import cprint
import requests

# Speak function using pywin32 module


def speak(audio):
    str = Dispatch("SAPI.spvoice")
    str.Speak(audio)


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        cprint('Good morning', 'magenta')
        speak("Good Morning!")

    elif hour >= 12 and hour < 18:
        cprint('Good Afternoon', 'magenta')
        speak("Good Afternoon!")

    else:
        cprint('Good evening', 'magenta')
        speak("Good Evening!")

    speak("I am Jarvis Sir. Please tell me how may I help you")


def take_commands():
    # Initialising microphone from speech recognization

    r = sr.Recognizer()
    with sr.Microphone() as source:
        cprint("Listening...", 'magenta')
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        cprint("Recognizing...", 'magenta')
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print("Say that again please...")
        return "None"
    return query


if __name__ == "__main__":
    wishMe()
    while True:
        user_command_list = []
        user = take_commands().lower()
        user_command_list.append(user)

        if 'open youtube' in user:
            webbrowser.open("youtube.com")

        elif 'open google' in user:
            webbrowser.open("google.com")

        elif 'play song' or 'song' in user:
            user = user.replace('Jarvis play','')
            print("Searching song")
            webbrowser.open(
                'https://www.youtube.com/results?search_query=' + user + 'song')
        elif 'time' or 'tell me the time' in user:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
        elif 'Jarvis tell me my history' or 'history' in user:
            for i in user_command_list:
                print(i, end=" ")
                speak(i)
        elif 'Jarvis tell me the weather in London' in user:
            api_address = 'http://api.openweathermap.org/data/2.5/weather?appid=0c42f7f6b53b244c78a418f4f181282a&q='
            user = user.replace(
                'Jarvis tell me the weather in', '')
            city = user
            url = api_address + city
            print(url)
            json_data = requests.get(url).json()
            format_add = json_data['weather']
            temp = json_data['main']
            for key, val in format_add[0].items():
                if key == 'description':
                    print('wheather is', val)
                    speak('wheather is', val)
            for key, val in temp.items():
                if key == 'temp':
                    print('temperature is', val)
                    speak('temperature is', val)
                    
        elif 'Jarvis tell me my history' or 'history' in user:
            for i in user_command_list:
                print(i, end=" ")
                speak(i)

        elif 'Jarvis spell' in user:
            user = user.replace('Jarvis spell', '')
            speak(user)
        

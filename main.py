from gtts import gTTS
import speech_recognition as sr
import os
import re
import webbrowser
import requests
import win32com.client as wincl
import datetime 
import pywhatkit
import wikipedia
speak = wincl.Dispatch("SAPI.SpVoice")

def talkToMe(audio):
    
    print(audio)
    speak.Speak(audio)

def myCommand():   

    r = sr.Recognizer()

    with sr.Microphone() as source:
        speak_command = 'Please give some command to me...'
        talkToMe(speak_command)
        
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)

    try:
        command = r.recognize_google(audio).lower()
        talkToMe('You said: ' + command + '\n')

    
    except sr.UnknownValueError:
        talkToMe('Your last command couldn\'t be heard ! I can understand commands like  open gmail, open website xyz.com and tell me a joke')
        
        command = myCommand()

    return command


def assistant(command):
    
    message = 'hello im alex , hi its really good to hear from you im here to help you with anything you want  like  open gmail, open website xyz.com ,book train tickets , tell me a joke and many more '
    if 'hello' in command:
        talkToMe(message)

    elif 'hi' in command:
        talkToMe(message)

    elif 'hey' in command:
        talkToMe(message)
    
    elif 'hello alex' in command:
        talkToMe(message)

    elif 'open gmail' in command:
        
        url = 'https://www.gmail.com/'
        webbrowser.open(url)
        talkToMe('Done!')

    elif 'open any game' in command:
        
        url = 'https://doodlecricket.github.io/#/'
        webbrowser.open(url)
        talkToMe(' start playing ')

    elif 'book train tickets' in command:
        
        url = 'https://www.irctc.co.in/nget/train-search'
        webbrowser.open(url)
        talkToMe('Done!')

    elif 'i want to shop' in command:
        
        url = 'https://www.amazon.in/'
        webbrowser.open(url)
        talkToMe('select your choice')

    elif 'i am hungry' in command:
        
        url = 'https://www.zomato.com/'
        webbrowser.open(url)
        talkToMe('order your favorite food')

    elif 'open website' in command:
        reg_ex = re.search('open website (.+)', command)
        if reg_ex:
            domain = reg_ex.group(1)
            url = 'https://www.' + domain
            webbrowser.open(url)
            talkToMe('Done!')
        else:
            pass

    elif 'open notepad' in command:
        os.startfile('C:\\Users\\rishi\\Desktop\\Alex\\alex.txt')
        talkToMe('Done for you!')

    elif 'time' in command:
        time = datetime.datetime.now().strftime('%I:%M %p')
        talkToMe('Current time is ' + time)

    elif 'open word' in command:
        os.startfile('C:\\Users\\rishi\\Desktop\\Alex\\alex.docx')
        talkToMe('Done for you!')
    
    elif 'open calculator' in command:
        os.startfile('C:\\Windows\\System32\\calc.exe')
        talkToMe('Done for you!')

    elif 'play' in command:
        song = command.replace('play', '')
        talkToMe('playing ' + song)
        pywhatkit.playonyt(song)

    elif 'who is' in command:
        person = command.replace('who is', '')
        info = wikipedia.summary(person, 2)
        talkToMe(info)

    elif 'how are you alex' in command:
        talkToMe('Just doing my things')

    elif 'tell me a joke' in command:
        res = requests.get(
                'https://icanhazdadjoke.com/',
                headers={"Accept":"application/json"}
                )
        if res.status_code == requests.codes.ok:
            talkToMe('Here is an awesome joke for you')
            talkToMe(str(res.json()['joke']))
        else:
            talkToMe('oops!I ran out of jokes')

    
    elif 'bye' in command:
        talkToMe('See you later Take care !')
        exit()
    elif 'shutdown system' in command:
        talkToMe("shutting down your system")
        os.system("shutdown /s /t 0")

    else:
            talkToMe('I don\'t know what you mean! I can understand commands like send email, open gmail, open website xyz.com and tell me a joke')

talkToMe('Hey, I am your laptop\'s virtual assistant alex and I am ready for your command !')


while True:
    assistant(myCommand())
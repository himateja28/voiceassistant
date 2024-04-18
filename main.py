import sys

import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import wikipedia
speaker = win32com.client.Dispatch("SAPI.SpVoice")
sites = [['youtube',"https://youtube.com"],
             ['instagram',"https://instagram.com"],['twitter',"https://twitter.com"],]
def takcommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold=0.5
        print("Listening........")
        try:
            audio = r.listen(source)
            query = r.recognize_google(audio, language="en-in")
            print(query)
            return query
        except Exception as e:
            return "Sorry I can not understand"
h = int(datetime.datetime.now().strftime("%H"))
if h<12:
    speaker.Speak("Good morning Teja sir")
elif h>=12 and h<16:
    speaker.Speak("Good afternoon Teja sir")
elif h>=16 and h<19:
    speaker.Speak("Good evening Teja sir")
else:
    speaker.Speak("Good night Teja sir ")
while True:
    sp = takcommand()
    speaker.Speak(sp)
    if "What the time now".lower() in sp.lower():
        ctime = datetime.datetime.now().strftime("%H:%M:%S")
        speaker.Speak(f"Time is {ctime} sir")
    elif 'wikipedia'.lower()in sp.lower():
        speaker.Speak('Searching Wikipedia...')
        sp = sp.replace("wikipedia", "")
        results = wikipedia.summary(sp, sentences=2)
        speaker.Speak("According to Wikipedia")
        print(results)
        speaker.Speak(results)
    elif 'shutdown':
        sys.exit()
    for site in sites:
        if site[0].lower() in sp.lower():
            webbrowser.open(site[1])
            speaker.Speak(f"Opening {site[0]}  sir ")


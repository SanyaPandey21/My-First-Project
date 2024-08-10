import ctypes
from PyQt5 import QtWidgets, uic
import datetime
import os
import shutil
import smtplib
import subprocess
import time
import webbrowser
import sys
import pyjokes
import pyttsx3
import requests
import speech_recognition as sr
import wikipedia
import winshell
import wolframalpha
from bs4 import BeautifulSoup
from clint.textui import progress
from ecapture import ecapture as ec
from finalpro import Ui_MainWindow
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QDate, Qt, QTime, QTimer
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtGui import QMovie
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from PyQt5.QtCore import QThread

engine = pyttsx3.init('sapi5')

voices = engine.getProperty('voices')

engine.setProperty('voice', voices[1].id)

def speak(audio):

    engine.say(audio)

    engine.runAndWait()

def wishMe():

    hour = int(datetime.datetime.now().hour)

    if hour>= 0 and hour<12:

        speak("Good Morning Sir !")

  

    elif hour>= 12 and hour<18:

        speak("Good Afternoon Sir !")   

  

    else:

        speak("Good Evening Sir !")  

  

    assname =("Jarvis 1 point o")

    speak("I am your Assistant")

    speak(assname)



class MainThread(QThread):
    def __init__(self):
        super(MainThread,self).__init__()
                
    def run(self):
        self.Taskexecution()

    def takeCommand(self):

        r = sr.Recognizer()

        with sr.Microphone() as source:

            print("Listening...")

            r.pause_threshold = 1

            audio = r.listen(source)

        try:

            print("Recognizing...")    

            query = r.recognize_google(audio, language ='en-in')

            print(f"User said: {query}\n")

        except Exception as e:

            print(e)    

            print("Unable to Recognize your voice.")  

            return "None"

        return query

    def Taskexecution(self):

        #clear = lambda: os.system('cls')

        

        # This Function will clean any

        # command before execution of this python file

        #clear()
        engine = pyttsx3.init('sapi5')

        voices = engine.getProperty('voices')

        engine.setProperty('voice', voices[1].id)

        def speak(audio):

            engine.say(audio)

            engine.runAndWait()
        
        wishMe()

        speak("What should i call you sir")

        uname = self.takeCommand()

        speak("Welcome Mister")

        speak(uname)

        columns = shutil.get_terminal_size().columns

     

        print("#####################".center(columns))

        print("Welcome Mr.", uname.center(columns))

        print("#####################".center(columns))

     

        speak("How can i Help you, Sir")


        

        while True:

            

            query = self.takeCommand().lower()

            

            # All the commands said by user will be 

            # stored here in 'query' and will be

            # converted to lower case for easily 

            # recognition of command

            if 'wikipedia' in query:

                speak('Searching Wikipedia...')
                query = query.replace("wikipedia", "")
                results = wikipedia.summary(query, sentences=2) 
                speak("According to Wikipedia")
                print(results)
                speak(results)
    

            elif 'open youtube' in query:

                speak("Here you go to Youtube\n")

                webbrowser.open("youtube.com")
    

            elif 'open google' in query:

                speak("Here you go to Google\n")

                webbrowser.open("google.com")
    

            elif 'open stackoverflow' in query:

                speak("Here you go to Stack Over flow.Happy coding")

                webbrowser.open("stackoverflow.com")   
    

            elif 'play music' in query or "play song" in query:

                speak("Here you go with music")

                # music_dir = "G:\\Song"

                music_dir = "C:\\Users\\aparn\\Music\\music"

                songs = os.listdir(music_dir)

                print(songs)    

                random = os.startfile(os.path.join(music_dir, songs[0]))

            elif 'open calculator' in query:
                oreo = r"C:\\Windows\\System32\\calc.exe"
                os.startfile(oreo)
                


            elif 'open notepad' in query:
                no = r"C:\\Windows\\System32\\Notepad.exe"     
                os.startfile(no)

            elif 'the time' in query:

                strTime = datetime.datetime.now().strftime("%H:%M:%S")    
                speak(f"Sir, the time is {strTime}")
    

            elif 'open opera' in query:

                codePath = r"C:\\Users\\aparn\AppData\\Local\\Programs\\Microsoft VS Code"

                os.startfile(codePath)
    

            elif 'how are you' in query:

                speak("I am fine, Thank you")

                speak("How are you, Sir")
    

            elif 'fine' in query or "good" in query:

                speak("It's good to know that your fine")
    

            elif "change my name to" in query:

                query = query.replace("change my name to", self.takeCommand())

                assname = query
    

            elif "change name" in query:

                speak("What would you like to call me, Sir ")

                assname = self.takeCommand()

                speak("Thanks for naming me")
    

            elif "what's your name" in query or "What is your name" in query:

                speak("My friends call me")

                speak(assname)

                print("My friends call me", assname)
    

            elif 'exit' in query:

                speak("Thanks for giving me your time")

                exit()
    

            elif "who made you" in query or "who created you"  in query: 

                speak("I have been created by my team.")

                

            elif 'joke' in query:

                speak(pyjokes.get_joke())

            elif 'search' in query or 'play' in query:
                statement = statement.replace("search", "")
                webbrowser.open_new_tab(statement)
                time.sleep(5)

            elif "who i am" in query:

                speak("If you talk then definitely your human.")
    

            elif "why you came to world" in query:

                speak("Thanks to Aman,Manas,Aparna. further It's a secret")
    

            elif 'power point presentation' in query:

                speak("opening Power Point presentation")

                power = r"C:\\Users\\aparn\\OneDrive\\Desktop\\jervis.pptx"

                os.startfile(power)
    

            elif "who are you" in query:

                speak("I am your virtual assistant created by Aman,Manas,Aparna")
    

            elif 'reason for you' in query:

                speak("I was created as a Major  project by Aman,Manas,Aparna ")
    

            elif 'change background' in query:

                ctypes.windll.user32.SystemParametersInfoW(20, 

                                                        0, 

                                                        "C:\\Windows\\Web\\Wallpaper\\Spotlight\\img14.jpg",

                                                        0)

                speak("Background changed successfully")
    

            elif 'news' in query:

                '''news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/city/mangalore")
                speak('Here are some headlines from the Times of India, Happy reading')'''
                    #time.sleep(6)
                query_params = {
                    "source": "bbc-news",
                    "sortBy": "top",
                    "apiKey": "815e9ab2e4724949b2bfd342861c4a06"
                    }
                main_url = " https://newsapi.org/v1/articles"
                res = requests.get(main_url, params=query_params)
                open_bbc_page = res.json()
                article = open_bbc_page["articles"]
                results = []
                for ar in article:
                    results.append(ar["title"])
                for i in range(len(results)):
                    print(i + 1, results[i])
                from win32com.client import Dispatch
                speak = Dispatch("SAPI.Spvoice")
                speak.Speak(results)
    
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

                speak("for how much time you want to stop jarvis from listening commands")

                a = int(self.takeCommand())

                time.sleep(a)

                print(a)
    

            elif "where is" in query:

                self.query = query.replace("where is", "")

                location = query

                speak("User asked to Locate")

                speak(location)

                webbrowser.open("https://www.google.com/maps/@26.9182799,80.9539467,15z" + location + "")
    

            elif "open camera" in query or "take a photo" in query:

                ec.capture(0, "Jarvis Camera ", "img.jpg")
            
            elif "weather" in query:
                import requests
                from bs4 import BeautifulSoup
                headers = {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
                def weather(city):
                    city = city.replace(" ", "+")
                    res = requests.get(
                        f'https://www.google.com/search?q={city}&oq={city}&aqs=chrome.0.35i39l2j0l4j46j69i60.6128j1j7&sourceid=chrome&ie=UTF-8', headers=headers)
                    print("Searching...\n")
                    soup = BeautifulSoup(res.text, 'html.parser')
                    location = soup.select('#wob_loc')[0].getText().strip()
                    time = soup.select('#wob_dts')[0].getText().strip()
                    info = soup.select('#wob_dc')[0].getText().strip()
                    weather = soup.select('#wob_tm')[0].getText().strip()
                    speak(location)
                    speak("Time is")
                    speak(time)
                    speak("other information")
                    speak(info)
                    speak("temperature is")
                    speak(weather+"°C")
                speak("What is city name?")
                city = self.takeCommand()
                city = city+" weather"
                weather(city)
                speak("Have a Nice Day:)")

    

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

                note = self.takeCommand()

                file = open('C:\\Users\\aparn\\OneDrive\\Desktop\\Voice.txt', 'w')

                speak("Sir, Should i include date and time")

                snfm = self.takeCommand()

                if 'yes' in snfm or 'sure' in snfm:

                    strTime = datetime.datetime.now().strftime("%H:%M:%S")  

                    file.write(strTime)

                    file.write(" :- ")

                    file.write(note)

                else:

                    file.write(note)

            

            elif "show note" in query:

                speak("Showing Notes")

                file = open("C:\\Users\\aparn\\OneDrive\\Desktop\\Voice.txt", "r") 

                print(file.read())

                speak(file.read(6))
    
            # NPPR9-FWDCX-D2C8J-H872K-2YT43

            elif "jarvis" in query:

                wishMe()

                speak("Jarvis 1 point o in your service Mister")

                speak(assname)
    
            elif "wikipedia" in query:

                webbrowser.open("wikipedia.com")
    

            elif "Good Morning" in query:

                speak("A warm" +query)

                speak("How are you Mister")

                speak(assname)
    

            # most asked question from google Assistant


            elif "how are you" in query:

                speak("I'm fine, glad you me that")

startExecution = MainThread()



class Main(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.startTask)
        self.ui.pushButton_2.clicked.connect(self.close)
        
        
    def startTask(self):
        self.ui.movie1 = QMovie("C:/Users/aparn/OneDrive/Desktop/robot.gif")
        self.ui.jarvisUi.setMovie(self.ui.movie1)
        self.ui.movie1.start()
        
        self.ui.movie2 = QMovie("C:/Users/aparn/OneDrive/Desktop/initial.gif")
        self.ui.label_2.setMovie(self.ui.movie2)
        self.ui.movie2.start()
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.showTime)
        self.timer.start(1000)
        
        # assuming startExecution is a function defined elsewhere
        startExecution.start()  

    def showTime(self):
        current_time = QTime.currentTime()
        current_date = QDate.currentDate()
        label_time = current_time.toString('hh:mm:ss')
        label_date = current_date.toString(Qt.ISODate)
        self.ui.textBrowser.setText(label_date)
        self.ui.textBrowser_2.setText(label_time)
        
        
        
    def closeEvent(self, event):
        self.timer.stop()
        super().closeEvent(event)


 #client = wolframalpha.Client("4V2RGY-TLLRW4KXLQ")
        # elif "" in query:
            # Command go here
            # For adding more commands

app=QApplication(sys.argv)
jarvis = uic.loadUi("C:/Users/aparn/OneDrive/Desktop/JarvisGui/jarvisUi.ui")
jarvis=Main()
jarvis.show()
exit(app.exec_())




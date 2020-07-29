import pyttsx3 as ps
import datetime as dt
import speech_recognition as sr
import wikipedia as wiki
import webbrowser as wb
import os
import random
import time
import smtplib
from selenium import webdriver
import re
from mutagen.mp3 import MP3

engine=ps.init('sapi5')
voices=engine.getProperty('voices')
#print(voices[1].id) #to get the voices from window
engine.setProperty('voice',voices[1].id)
chrome_path="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
wb.register('google-chrome', None,wb.BackgroundBrowser(chrome_path))


def welcome():
    print("||||||||||      |||||||||||||      |||        |||           |||||||||")
    time.sleep(0.5)
    print("|||    |||           |||             |||    |||             |||   |||")
    time.sleep(0.5)
    print("|||    |||           |||               ||||||               |||   |||")
    time.sleep(0.5)
    print("||||||||||           |||                 ||                 |||   |||")
    time.sleep(0.5)
    print("||| ||               |||                 ||                 |||||||||")
    time.sleep(0.5)
    print("|||   ||             |||                 ||                 |||   |||")
    time.sleep(0.5)
    print("|||     ||           |||                 ||                 |||   |||")
    time.sleep(0.5)
    print("|||      ||     |||||||||||||            ||                 |||   |||")
    time.sleep(1)
    print("\033[1m", "\n                  Your personel voice assistant", "\033[0m")








def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour= int(dt.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning Mate! \n Have a good day.")

    elif hour>=12 and hour<=18:
        speak("Good Afternoon Good Afternoon Jay! \n Hope,you are having a good day.")

    else:
        speak("Good Evening! \n Hope you had a great day.")

    speak("I am Riya. The personnel voice assistant created by Jay Khandelwal. Please tell me How may I help you?")


def takeCommand():
    #takes microphone input from user and return string output

    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("I'm listening....")
        r.pause_threshold=1
        audio=r.listen(source)

    try:
        print("I'm Recognizing....")
        query= r.recognize_google(audio, language='en-in')
        print(f"You just said: {query}\n")

    except Exception as e:
        print(e)

        print("Unable to recognize! Can you please say it again...")
        return "None"
    return query

def sendEmail(to,content):
    server=smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    mail='jayk181099@gmail.com'
    server.login(mail,'Dev_jay18')
    server.sendmail(mail,to,content)
    server.close()



def search_google(query):
    chrome_driver = webdriver.Chrome(executable_path="C:\Drivers\chrome\chromedriver.exe")
    chrome_driver.get(url="https://www.google.com")
    search_query=chrome_driver.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input')
    search_query.click()
    search_query.send_keys(query)
    submit=chrome_driver.find_element_by_name('btnK')
    #submit=self.driver.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[3]/center/input[1]').click()
    submit.click()

def play_youtube(query):
    chrome_driver = webdriver.Chrome(executable_path="C:\Drivers\chrome\chromedriver.exe")
    chrome_driver.get(url="https://www.youtube.com/results?search_query="+query)
    video_play=chrome_driver.find_element_by_xpath('//*[@id="img"]')
    video_play.click()
    #time.sleep(90)

def general_chat(query):
    if query in question:
        answer=question[query]
        speak(answer)

def sleep(sleep_time):
    time.sleep(sleep_time)
    speak("Again available to serve you, Jay Sir")

def close_chrome():
    os.system('taskkill /f /im chrome.exe')


if __name__ == '__main__':
    welcome()
    wishMe()

    #load general chat/conversion data
    question = {}
    with open("dict.txt") as f:
        for line in f:
            (key, val) = line.split('-')
            question[(key)] = val

    while True:
        query=takeCommand().lower()

        #logics for executing tasks based on query
        if 'search' in query:
            speak('Searching your query...')
            query=query.replace("search","")
            search_google(query)

        elif 'open youtube' in query:
            speak('We will getting that from youtube soon')
            query=query.replace("open youtube ","")
            play_youtube(query)

        elif 'tell me about' in query:
            speak('We will let you know soon...')
            query = query.replace("tell me about", "")
            results = wiki.summary(query, sentences=2)
            speak("According to our research")
            speak(results)
            print(results)

        elif 'website open ' in query:
            position_open=query.find('open')
            position_open+=5
            ur=query[position_open:]+".com"
            urL=ur.replace(" ", "")
            wb.get('google-chrome').open(urL)

        elif 'close chrome' in query:
            speak("closing chrome")
            close_chrome()

        elif 'play music' in query:
            music_dir=r'C:\Data\Songs'
            songs=os.listdir(music_dir)
            song_number=random.randint(0,len(songs)-1)
            #audio=MP3(songs[song_number])
            #song_length=audio.info.length
            os.startfile(os.path.join(music_dir,songs[song_number]))
            time.sleep(30)

        elif 'play video' in query:
            video_dir = r'C:\Data\Mobile'
            video = os.listdir(video_dir)
            video_number = random.randint(0, len(video) - 1)
            os.startfile(os.path.join(video_dir, video[video_number]))
            time.sleep(30)

        elif 'the time' in query:
            strTime=dt.datetime.now().strftime("%H:%M:%S")
            speak(f"Jay, The time is {strTime}")

        elif 'open world' in query:
            word_path=r"C:\Softwares\word.lnk"
            os.startfile(word_path)

        elif 'open powerpoint' in query:
            point_path=r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            os.startfile(point_path)

        elif 'open excel' in query:
            excel_path=r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            os.startfile(excel_path)

        elif 'open python editor' in query:
            pycharm_path=r"C:\Users\Dev_jay\AppData\Local\JetBrains\Toolbox\apps\PyCharm-P\ch-0\201.7846.77\bin\pycharm64.exe"
            os.startfile(pycharm_path)




        elif 'hold' in query:
            temp = re.findall(r'\d+', query)
            res = list(map(int, temp))
            sleep_time=res[0]
            sleep(sleep_time)


        elif 'send email' in query:
            try:
                speak("What should I say?")
                content=takeCommand()
                to="jay.khandelwal_bca16@gla.ac.in"
                sendEmail(to,content)
                speak("Email has been sent successfully!")
            except Exception as e:
                print(e)
                speak("Sorry, Unable to send Email")

        else:
            general_chat(query)




import speech_recognition as sr
import os
import win32com.client
import webbrowser
import datetime
import time
import smtplib
import openai
from config import apikey
from credentials import myemailid, password

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatStr = ""
# TODO: CHAT METHOD
def chat(query):
    global chatStr
    openai.api_key = apikey
    chatStr += f"Saurabh: {query} Jarvis: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=query,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    print("Jarvis :" + response["choices"][0]["text"])
    say(response["choices"][0]["text"])
    chatStr += response["choices"][0]["text"]
    return response["choices"][0]["text"]



# TODO: AI METHOD
def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for prompt: {prompt} \n *******************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    #print(response["choices"][0]["text"])
    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(prompt.split('intelligence')[1:])}.txt", "w") as f:
        f.write(text)
    say("Output generated and saved in a Open AI folder sir...")

# TODO: MAIL METHOD
def sendMail(emailid, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(myemailid, password)
    server.sendmail(myemailid, emailid, content)
    server.close();


# TODO: SPEAKING METHOD
def say(text):
    speaker.Speak(text)

# TODO: INPUT COMMAND
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source, 5, 10)
        try:
            print("recognizing")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occured Sorry sir..."




if __name__ == '__main__':
    print('Pycharm')
    say("Hello I am Saurabh's assistant. my name is Jarvis")
    while True:
        print("Listening...")
        query = takeCommand()

        #todo: for websites
        sites = [["gmail", "https://mail.google.com/mail/u/0/#inbox"], ["youtube", "https://youtube.com"], ["gfg", "https://practice.geeksforgeeks.org"], ["lead code", "https://leetcode.com/problemset/all/"], ["wikipedia", "https://wikipedia.com"],["chrome", "https://google.com"],["Graphic Era Hill University Website", "https://www.gehu.ac.in/"], ["Graphic Era Hill University ERP", "https://student.gehu.ac.in/Account/Cyborg_StudentMenu"]]

        for site in sites:
            if f" {site[0]}".lower() in query.lower():
                say(f"opening {site[0]} sir...")
                webbrowser.open(site[1])

        #todo: for system file/folders
        lists = [["music list", "D:\my music"], ["play music", "D:\my music\Make You Mine.mp3"], ["play song", "D:\my music\Make You Mine.mp3"] ]

        if "open".lower() in query.lower():
            for list in lists:
                if f"{list[0]}".lower() in query.lower():
                    say(f"opening {list[0]} sir...")
                    os.startfile(list[1])


        #todo: for system apps
        apps = [["camera", "start microsoft.windows.camera:"], ["calculator", "calc"], ["notepad", "notepad"]]

        if "open".lower() in query.lower():
            for app in apps:
                if f"{app[0]}".lower() in query.lower():
                    say(f"opening {app[0]} sir...")
                    os.system(app[1])


        # todo : for killing processes
        processes = [["camera", "taskkill /im WindowsCamera.exe /f"], ["google tabs", "taskkill /im chrome.exe /f"], ["file explorer", "taskkill /im explorer.exe /f "], ["music", "taskkill /im vlc.exe"]]

        if "close".lower() in query.lower():
            for process in processes:
                if f"{process[0]}".lower() in query.lower():
                    say(f"closing {process[0]} sir...")
                    os.system(process[1])




        #todo: for system commands
        commands = [["shutdown", "shutdown /s"], ["restart", "shutdown /r"], ["log of", "shutdown /l"]]

        for command in commands:
            if f"{command[0]}".lower() in query.lower():
                say(f"{command[0]} the system sir...")
                os.system(command[1])

        #todo: for sending emails
        emailids = [["saurabh", "bhattsourav23@gmail.com"], ["rishabh kukreti", "kukretirishabh07@gmail.com"]]

        if "send a mail to" in query:
            for name in emailids:
                if f"{name[0]}".lower() in query.lower():
                    try:

                        say("say what you want to send in mail")
                        time.sleep(3)
                        say("say now sir...")
                        content = takeCommand()
                        print(content+"\n")
                        sendMail(name[1], content)
                        say("email successfully sended sir...")
                    except Exception as e:
                        print(e)
                        say("Sorry sir I am not able to send this email right now please try again")



        elif "the time" in query.lower():
            strftime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Sir the time is {strftime}")

        elif "the date" in query.lower():
            strfdate = datetime.datetime.today().strftime('%Y-%m-%d')
            say(f"Sir the time is {strfdate}")


        elif "vs code" in query:
            path = "C:\\Users\\hp\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            say("opening vs code sir")
            os.startfile(path)

        elif "reset chat".lower() in query.lower():
            say("chat reset Sir...")
            chatStr = ""

        elif "using artificial intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "chatting".lower() in query.lower():
            say("chatting started Sir...")
            #query = takeCommand()
            while True:
                print("chatting")
                query = takeCommand()
                if "stop chatting".lower() in query.lower():
                    say("chat stopped Sir...")
                    break
                chat(query)

        elif "stop working" in query.lower():
            say("ok sir going to stop")
            quit()


        #else:    say("Sorry sir I am not able to understand your query please try again")






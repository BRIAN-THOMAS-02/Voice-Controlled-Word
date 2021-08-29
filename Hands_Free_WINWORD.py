import os
import docx
import speech_recognition as sr
import playsound
import time
import pyaudio
import pyttsx3
import time
from datetime import datetime
from datetime import date
import pyautogui
import keyword

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
print("")
# print(voices[0].id)
engine.setProperty('voice', voices[1].id)
newVoiceRate = 165
engine.setProperty('rate', newVoiceRate)

docs = docx.Document()


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def get_audio():
    import speech_recognition as sr
    rObject = sr.Recognizer()
    audio = ''

    with sr.Microphone() as source:
        print(" ")
        print("Listening...")
        audio = rObject.listen(source, phrase_time_limit=3)

        try:
            text = rObject.recognize_google(audio, language='en-us')
            print("You said... :", str.capitalize(text))
            print(" ")
            return text

        except:
            print(" ")
            print("Could not understand you Brian, PLease try again !")
            speak("Could not understand you Brian, PLease try again !")
            print(" ")
            return 0


def write_content():
    while True:
        keyword = get_audio()
        if "continue" in str(keyword).lower():
            while True:
                global writing
                writing1 = get_audio1()
                writing = str(writing1)

                #xyz = str(writing)
                #pyautogui.write(xyz)
                #pyautogui.press('space')

                if "next line" in str(writing1).lower():
                    pyautogui.press("enter")

                if "bold" in str(writing1).lower():
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("b")
                    pyautogui.keyUp("ctrl")

                if "italics" in str(writing1).lower():
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("i")
                    pyautogui.keyUp("ctrl")

                if "underline" in str(writing1).lower():
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("u")
                    pyautogui.keyUp("ctrl")

                if "left align" in str(writing1).lower():
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("l")
                    pyautogui.keyUp("ctrl")

                if "right align" in str(writing1).lower():
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("r")
                    pyautogui.keyUp("ctrl")

                if "center align" in str(writing1).lower() or "centre align" in str(writing1).lower():
                    pyautogui.keyDown("ctrl")
                    pyautogui.press("e")
                    pyautogui.keyUp("ctrl")

                sd = (writing.replace("Centre align", " ").replace("next line", " ").replace("left align", " ").replace("right align", " ").replace("bold", " ").replace("underline", " ").replace("italics", " "))
                pyautogui.write(sd)
                pyautogui.press('space')

                continue

                if "stop" in str(writing).lower():
                    global seconds5
                    seconds1 = (datetime.now() - start).seconds
                    seconds5 = int(seconds1)
                break
        break


def get_audio1():
    import speech_recognition as sr
    rObject = sr.Recognizer()
    audio = ''

    with sr.Microphone() as source:
        print(" ")
        print("Listening...")
        audio = rObject.listen(source, phrase_time_limit=6)

        try:
            text = rObject.recognize_google(audio, language='en-us')
            print("You said... :", str.capitalize(text))
            print(" ")
            return text

        except:
            print(" ")
            print("PLease try again...")
            speak("PLease try again...")
            print(" ")
            return 0


def secs_counter():
    from datetime import datetime as dt
    try:
        input1 = raw_input
    except:
        pass
    e = input1("Press Enter")
    start = dt.now()
    e = input1("Press Enter again")
    seconds = (dt.now() - start).seconds
    print("You took {} seconds between Enter actions".format(seconds))


if __name__ == "__main__":

    word = get_audio()
    if "open" in str(word).lower():
        print("Opening Microsoft Word...")
        speak("Opening Microsoft Word...")
        os.system("start Hands_free_test.docx")
        time.sleep(2)
        write_content1 = get_audio()
        if "start writing" in str(write_content1).lower():
            print("Say Continue to Start Writing and Stop Typing to Stop Writing")
            speak("Say Continue to Start Writing and Stop Typing to Stop Writing")
            write_content()

#secs_counter()

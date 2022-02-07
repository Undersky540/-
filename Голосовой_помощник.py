import speech_recognition
import os
import random
import pyttsx3
from win32com.client import constants
import win32com.client
import pythoncom


speaker = win32com.client.Dispatch("Sapi.SpVoice")

# устанавливаем скорость произношения от -10 до 10
speaker.Rate = 3
# устанавливаем громкость голоса от 0 до 100
speaker.Volume = 100

sr = speech_recognition.Recognizer()
sr.pause_threshold = 0.5
use_bot = True

def listen_command():
    try:
        with speech_recognition.Microphone() as mic:
            sr.adjust_for_ambient_noise(source=mic, duration=0.5)
            audio = sr.listen(source=mic)
            query = sr.recognize_google(audio_data=audio, language="ru-RU").lower()
        return query
    except speech_recognition.UnknownValueError:
        return ("Повторите...")


def greeting():
    # функция приветствия
    speaker.Speak("Приветствую Вас!")

def goodbye():
    # функция прощания
    speaker.Speak("Пока, зови еще!")
    print ("Пока, зови еще!")
    quit(0)

def creat_task():
    print ("Что добавить в список дел?")
    speaker.Speak("Что добавить в список дел?")
    print ("Каждое добавление в список дел выполняется отдельной командой: добавить задачу")
    speaker.Speak("Каждое добавление в список дел выполняется отдельной командой: добавить задачу")
   
             
    query = listen_command()

    with open("task-list.txt", "a", encoding='utf-8') as file: # добавляем "а" для добавления в файл, иначе будет перезаписываться а не добавлялться
        file.write(f"{query}\n")
    print (f"Задача {query} добавлена в task-list.txt")
    speaker.Speak (f"Задача {query} добавлена в task-list.txt")

def play_sound():
    pass
def welcome():
    print("Приветствую Вас! Я Ваш голосовой помощник.")
    speaker.Speak("Приветствую Вас! Я Ваш голосовой помощник.")
    print ("Я понимаю команды: 1) добавить задачу, 2) привет 3) пока")
    speaker.Speak("Я понимаю команды: добавить задачу, привет, пока")    
    print ("Слушаю команду...")
    speaker.Speak("Слушаю команду...")
def main():
    welcome()
    while use_bot == True:
        query = listen_command()
        if query == "привет":
            print(greeting())
        elif query == "добавить задачу":
            print(creat_task())
        elif query == "пока":
            goodbye()
        elif query == "спасибо":
            print("Пожалуйста, человек")
            speaker.Speak("Пожалуйста, человек")
        else:
            speaker.Speak("Я не поняла, скажите еще раз")
            print("Я не поняла, повтори")

if __name__=="__main__":
    main()

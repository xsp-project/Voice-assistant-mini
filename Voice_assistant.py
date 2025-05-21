import tkinter as tk
import speech_recognition as sr
import pyttsx3 as tts
import pywhatkit as kit
from plyer import notification as notify
from datetime import datetime
import os
import io
import sys

r = sr.Recognizer()
engine = tts.init()

def speak_text(command):
    engine.say(command)
    engine.runAndWait()

def start_listening():
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=2)
        display_text.set("Listening...")
        try:
            audio_text = r.listen(source, phrase_time_limit=5)
            text = r.recognize_google(audio_text).lower()
            display_text.set(f"You said: {text}")
            process_command(text)
        except sr.UnknownValueError:
            display_text.set("Sorry, I did not understand that.")
            speak_text("Sorry, I did not understand that.")
        except sr.RequestError as e:
            display_text.set("Could not request results; check your internet connection.")
            speak_text("Could not request results; check your internet connection.")

def process_command(text):
    if "hello" in text:
        response = "Hello! How can I assist you today?"
        display_text.set(response)
        speak_text(response)

    elif "goodbye" in text:
        response = "Goodbye! Have a great day!"
        display_text.set(response)
        speak_text(response)
        window.quit()

    elif "on youtube" in text or "play" in text:
        response = "Playing on YouTube"
        display_text.set(response)
        speak_text(response)
        kit.playonyt(text.replace("on youtube", "").strip())

    elif "search" in text:
        query = text.replace("search", "").strip()
        response = f"Searching for {query}"
        display_text.set(response)
        speak_text(response)
        kit.search(query)
        
    elif "information" in text or "about" in text:
        response = "Fetching information for " + text.replace("information", "").strip()
        notify.notify(title="knowledge" , message ="Fetching...." , timeout = 4)
        speak_text(response)
        display_text.set(response)
        output = io.StringIO()
        sys.stdout = output
        query = text.replace("information", "").strip()
        info = kit.info(query, lines=3)
        sys.stdout = sys.__stdout__
        info = output.getvalue()
        print("\n :- "+info)
        speak_text(info)
        
    elif "shutdown" in text:
        kit.shutdown(time=30)
        notify.notify(title="Shut Down" , message ="Shuting Down in 30s. " , timeout = 5)
        response="Shutting your system down in 30 seconds"
        display_text.set(response)
        speak_text(response)
       
    elif "cancel shutdown" in text:
        kit.cancel_shutdown()
        response = "Shutdown cancelled"
        display_text.set(response)
        notify.notify(title="Shut Down" , message =response , timeout = 5)
        speak_text(response)
        
    elif "on whatsapp" in text or "send" in text:
        number = input("Enter the number: ")
        number = "+91" + number.strip()
        response = "Sending message on WhatsApp"
        display_text.set(response)
        print(response)
        notify.notify(title="WhatsApp" , message ="sending...." , timeout = 4)
        speak_text(response)
        msg = text.replace("send","").replace("on whatsapp", "").strip()
        kit.sendwhatmsg_instantly(number, msg)

    elif "date" in text:
        current_date = datetime.now().strftime("%A, %d %B %Y")
        response = f"Today's date is {current_date}"
        display_text.set(response)
        speak_text(response)

    elif "time" in text:
        current_time = datetime.now().strftime("%I:%M %p")
        response = f"The current time is {current_time}"
        display_text.set(response)
        speak_text(response)

    elif "open" in text:
        app_name = text.replace("open", "").strip()
        try:
            if "chrome" in app_name:
                os.startfile("C:/Program Files/Google/Chrome/Application/chrome.exe")
                response = "Opening Chrome"
            elif "notepad" in app_name:
                os.startfile("C:/Windows/System32/notepad.exe")
                response = "Opening Notepad"
            elif "file explorer" in app_name or "file" in app_name:
                os.startfile("D:")
                response = "Opening File Explorer"
            elif "excel" in app_name:
                os.startfile("C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE")
                response = "Opening Excel"
            elif "word" in app_name or "m s word" in app_name:
                os.startfile("C:/Program Files/Microsoft Office/root/Office16/WINWORD.EXE")
                response = "Opening Word"
            elif "powerpoint" in app_name or "power point" in app_name:
                os.startfile("C:/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE")
                response = "Opening PowerPoint"
            elif "calculator" in app_name:
                os.startfile("C:/Windows/System32/calc.exe")
                response = "Opening Calculator"
            elif "vlc" in app_name:
                os.startfile("C:/Program Files/VideoLAN/VLC/vlc.exe")
                response = "Opening VLC Media Player"
            else:
                site = f"https://{app_name}.com"
                os.system(f"start {site}")
                response = f"opening website {app_name}."
            display_text.set(response)
            speak_text(response)
            
        except Exception as e:
            response = f"Error opening {app_name}: {e}"
            display_text.set(response)
            speak_text(response)

    else:
        response = "Sorry, I can't perform that action."
        display_text.set(response)
        speak_text(response)
        
window = tk.Tk()
window.title("Voice Assistant")
window.geometry("600x400")
window.configure(bg="#2c3e57")

title_label = tk.Label(window, text="Voice Assistant", font=("Helvetica", 24, "bold"), bg="#2c3e50", fg="#ecf0f1")
title_label.pack(pady=20)

display_text = tk.StringVar()
display_text.set("Click 'Start Listening' to begin.")
display_label = tk.Label(window, textvariable=display_text, font=("Helvetica", 14), bg="#34495e", fg="#ecf0f1", wraplength=500, justify="center")
display_label.pack(pady=20, padx=20, fill="both")

listen_button = tk.Button(window, text="Start Listening", font=("Helvetica", 16), bg="#1abc9c", fg="#ffffff", command=start_listening)
listen_button.pack(pady=20)

exit_button = tk.Button(window, text="Exit", font=("Helvetica", 16), bg="#e74c3c", fg="#ffffff", command=window.quit)
exit_button.pack(pady=10)

window.mainloop()



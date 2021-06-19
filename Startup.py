import win32com.client as win
speak=win.Dispatch("SAPI.spVoice")
speak.rate= -1
speak.speak("Hi Abdelrahem")
speak.speak("good morning, How are you")
speak.speak("you are welcome to your computer")

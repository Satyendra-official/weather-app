import requests
import json
import win32com.client as wincom
import pyttsx3


# speak = wincom.Dispatch("SAPI.SpVoice")

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(text):
	engine.say(text)
	engine.runAndWait()

# text = "Python text-to-speech test. using win32com.client"
# speak(text)

city = input("Enter the name of the city : ")

url = f"http://api.weatherapi.com/v1/current.json?key=4509ed5c8a2a436f9f030901242503&q={city}"

r = requests.get(url)

# print(r.text)

wdic = json.loads(r.text)
region = (wdic["location"]["region"])
country = (wdic["location"]["country"])
wind_kph = (wdic["current"]["wind_kph"])
condition = (wdic["current"]["condition"]["text"])
deg_c = (wdic["current"]["temp_c"])
humidity = (wdic["current"]["humidity"])


temp = f"The temperature of {city} is {deg_c} degree celsius."
print(temp)
# speak.Speak(temp)
speak(temp)

reg = f"The City is in {region}."
print(reg)
speak(reg)
# speak.Speak(reg)

cntry = f"{city} is situated in {country}."
print(cntry)
speak(cntry)
# speak.Speak(cntry)

wind = f"The wind speed of {city} in Kilometer per hour is {wind_kph} Kilometer per hour."
print(wind)
speak(wind)
# speak.Speak(wind)

condi_n = f"The sky is {condition}."
print(condi_n)
speak(condi_n)
# speak.Speak(condi_n)

hmdt = f"The humidity at {city} is {humidity}."
print(hmdt)
speak(hmdt)
# speak.Speak(hmdt)

end = f"Thank you, have a nice day."
print(end)
speak(end)
# speak.Speak(end)

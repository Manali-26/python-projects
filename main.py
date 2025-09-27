import requests # makes HTTP requests
import json
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

city=input("enter the name the name of the city\n")

url=f"https://api.weatherapi.com/v1/current.json?key=c2a7e80d9e424ba5af0110322252709&q={city}"

r=requests.get(url)
# print(r.text)
wdic=json.loads(r.text)

temp=wdic["current"]["temp_c"]

print(f"The temperature in {city} is {temp}°C")

speaker.Speak(f" the tempreture in {city} is {temp} degree celsius")
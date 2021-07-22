from win32com.client import Dispatch
import requests
import json


def speak(st):
    say = Dispatch("SAPI.SpVoice")
    say.speak(st)


def news(st):
    url = f"https://newsapi.org/v2/top-headlines?country=in&category={st}&apiKey=99a2e524dc434d60856b017d63a79dff"
    result = requests.get(url)
    t = result.text
    final = json.loads(t)

    print(f"Top 5 news for {st} are:")
    for i in range(0, 5):
        speak(final['articles'][i]['title'])
        print(i+1, "->", final['articles'][i]['title'], "\n", "Time:", final['articles'][i]['publishedAt'])


print("Press 1 for business news")
print("Press 2 for entertainment news")
print("Press 3 for health news")
print("Press 4 for sports news")
print("Press 5 for science news: ")

x = int(input())
d = {1: 'business', 2: 'entertainment', 3: 'health', 4: 'sports', 5: 'science'}
news(d[x])

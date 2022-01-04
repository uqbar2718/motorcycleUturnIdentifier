# 지자체 목록 크롤링용

import requests
from bs4 import BeautifulSoup

url = 'https://ko.wikipedia.org/wiki/%EB%8C%80%ED%95%9C%EB%AF%BC%EA%B5%AD%EC%9D%98_%EA%B8%B0%EC%B4%88%EC%9E%90%EC%B9%98%EB%8B%A8%EC%B2%B4_%EB%AA%A9%EB%A1%9D'

response = requests.get(url)

list = []

html = response.text
soup = BeautifulSoup(html, 'html.parser')
# tbody = soup.select_one('tbody > tr > td > a',)
# print(tbody)

for location in soup.find_all('a'):
    list.append(location.text.strip())

print(list)
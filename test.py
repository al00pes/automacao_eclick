import requests
from bs4 import BeautifulSoup
from selenium import webdriver

response = requests.get('https://airbnb.com')

print(response.text)
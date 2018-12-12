# -*- coding: utf-8 -*-
__author__ = 'zhujingxiu'

from bs4 import BeautifulSoup

import requests


with open("https://baidu.com/") as htmldata:
    Soup = BeautifulSoup(htmldata,"lxml")
    print(Soup)

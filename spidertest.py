# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 17:08:00 2018

@author: Adu
"""

from requests_html import HTMLSession
from datetime import datetime

def HtmlDownloader(url):
    try:
        if url is None:
            return
        session = HTMLSession()
        r = session.get(url)
        return r.html
    except:
        return

def HtmlParser(url,html,path):
    date = {}
    postList = html.find('div.post')
    for post in postList:
        date['name'] = post.find('a.archive-title',first=True).text
        date['img_url'] = post.find('div.post-thumb',first=True).find('img',first=True).attrs['src']
        detail_url = post.find('span.read-more',first=True).find('a',first=True).attrs['href']
        date['detail_url'] = detail_url
        date['detail'] = HtmlDetailedParser(detail_url)[:12]
        date['time'] = datetime.now()
        with open(path,'a',encoding='utf-8') as f:
            f.write(str(date))
            f.write('\n')


def HtmlDetailedParser(url):
    html = HtmlDownloader(url)
    content = html.find('div.entry',first=True).text
    return content

def HtmlMian():
    path = 'F:\python_work\\text.txt'
    url = 'http://python.jobbole.com/all-posts/page/1/'
    html = HtmlDownloader(url)
    HtmlParser(url, html, path)

HtmlMian()
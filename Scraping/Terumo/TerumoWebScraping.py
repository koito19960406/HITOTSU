#!/usr/bin/env python
# coding: utf-8

# # Import Libraries 
# Use requests + beautifulsoup

# In[5]:


import requests
import csv
from bs4 import BeautifulSoup


# # create a function

# In[33]:


def scrapeTerumo(year,month,day):
    year=str(year)
    if month<10:
        month='0'+str(month)
    else:
        month=str(month)
    if day<10:
        month='0'+str(day)
    else:
        day=str(day)
    input_date=year+'年'+month+'月'+day+'日'
    # check if the URL is working
    url = 'https://www.terumo.co.jp/medical/news/equipment.html'
    result = requests.get(url)
    if result.status_code == 200:
        src=result.content
        soup=BeautifulSoup(src,"lxml")
        # get a list of news
        rows=soup.find('ul',attrs={'class':'newsList'}).find_all('li')
        # with open(input_date+company_name+'.csv','w') as file:
            # fieldnames=['日付','新着記事タイトル','新着記事URL']
            # writer=csv.DictWriter(file,fieldnames=fieldnames)
            # writer.writeheader()
            # goes through the list of news to get date, title, and URL
        news_list=[]
        for row in rows:
            date=row.find('dt').text.strip()
            # if the date matches the input date, store info in CSV
            if date==input_date:
                news_url=row.find('a').get('href')
                title=row.find('a').text.strip()
                news_list.append([date,title,news_url])
                # writer.writerow({
                # '日付':date,
                # '新着記事タイトル':title,
                # '新着記事URL':news_url
                # })
        return news_list


# # Main Part

# In[35]:


if __name__=='__main__':
    year=2020
    month=7
    day=31
    list=scrapeTerumo(year,month,day)
    print(list)


# In[ ]:





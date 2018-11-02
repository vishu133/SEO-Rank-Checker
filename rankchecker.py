# -*- coding: utf-8 -*-
"""
Created on Mon Nov 27 16:47:18 2017

@author: vish
"""

#rank checker
from bs4 import BeautifulSoup
from xlrd import open_workbook
import urllib3
import re
import time
import random
import datetime
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

"""adding user agent minimizes the chance of getting blocked"""
try:
    t0 = time.time()
    browsers = [' Mozilla/5.0 (Windows NT 6.1; WOW64; rv:12.0) Gecko/20100101 Firefox/12.0',
                'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36',
                'Mozilla/5.0 (Windows; U; Windows NT 6.1; x64; fr; rv:1.9.2.13) Gecko/20101203 Firebird/3.6.13',
                'Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; AS; rv:11.0) like Gecko',
                'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.3a) Gecko/20021207 Phoenix/0.5']
    
    
    
    path = input('enter excel file path where you have listed keywords ')
    book = open_workbook(path+'.xlsx')
    sheet = book.sheet_by_index(0)
    keywordsList = []
    
    for row in range(1,sheet.nrows): #start from 1, to leave out row 0
        keywordsList.append(str(sheet.cell(row, 0).value)) #extract from first col
        
    for keys in keywordsList:
        print(keywordsList.index(keys)+1,keys)
    webcheck = input("enter website ") #website whos keyword rankings you want to check
    
    
    
    filepath = path +'.csv'
    f = open(filepath,"w",encoding="utf-8")
    
    now = datetime.datetime.now()
    date = now.strftime("%d-%m-%Y %H:%M")
    
    headers = "keyword,rank,url,SERP title,Actual Title,ads present?,keywords present on landing page,current date: "+date+"\n"
    f.write(headers)    
    
    for keywordr in keywordsList:
        user_agent = {'user-agent': browsers[random.randint(0,len(browsers)-1)]}
        keyword = keywordr.replace(" ","+")
        url = "https://www.google.co.in/search?q=" + keyword + "&num=100"
        confuser = "https://www.google.co.in/search?q="+str(random.randint(0,100))
        
        
        http = urllib3.PoolManager(10, headers = user_agent, maxsize=10, block=True)
        response = http.request('GET', url)
        print("status: "+ str(response.status))
        
        soup = BeautifulSoup(response.data,'html.parser')
        containers = soup.findAll(lambda tag: tag.name == 'div' and tag.get('class') == ['g'])
        ads = soup.findAll('cite',{'class':'_WGk'})
        
        lists = []
        listTitle = []
        
    
    
        for container in containers: # creates list of 100 websites for the keyword
            try:
                website = (container.a['href'])
                title = (container.a.text)
                websiteRegex = re.compile(r'&.*',re.DOTALL)
                site = websiteRegex.sub('',website)
                
                
                site = site.replace('/url?q=','')
                lists.append(site)
                listTitle.append(title)
                
    
            except:
                continue
                
        for i in lists:
            if webcheck in i:
                
                if ads:
                    ad = 'yes'
                else:
                    ad = 'no'
                    
                rank = str(lists.index(i)+1)
                onpage = http.request('GET', i)
                soups = BeautifulSoup(onpage.data,'html.parser')
                PageTitle = soups.find('title')
                PageTitle = PageTitle.text
                keywordcount = str(soups.get_text().strip().lower().count(keywordr))
                f.write(keywordr + "," + rank + "," + i + "," + listTitle[lists.index(i)].replace(",","|") + "," + PageTitle.replace(",","-") + "," + ad + "," + keywordcount + "\n")
                print(keywordr,lists.index(i)+1)
                break
        try:
            if lists.index(i)+1 == 100:
                print("keyword ranks 100/above 100")
                f.write(keywordr + "," + "100" + "," + "none" + "," + "none" + "," + "none" + "," + "none" + "\n")    
        except:
            break
        time.sleep(random.uniform(8, 15))
    
    f.close()
    t1 = time.time()
    total = int((t1-t0)/60)
    print("file saved")
    print("time it took "+str(total)+" min")
    input("Your csv file has been saved in path "+path+".csv Press enter key to close the program...")
except Exception as e:
    f.close()
    print("file saved")
    print("system shut abruptly.")
    print(e)
    
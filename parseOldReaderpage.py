# import requests
from bs4 import BeautifulSoup
#import os

#Adjust two filenames below before running this script!!
htmlFile = open('oldreaderncscupdate20201104.html')
xmlFile = open('oldreaderexport20201104.xml','w')

#cwd = os.getcwd()
# print the current directory 
#print("Current working directory is:", cwd)


toParse = BeautifulSoup(htmlFile, "html.parser")
xmlFile.write('<?xml version="1.0" encoding="UTF-8"?>\n\n<rss version="2.0">\n\n<channel>\n')


for tag in toParse.find_all('h3'):
    #xmlFile.write("<item>\n" + str(tag) + "\n")
    subtag = tag.find_next('a')
    subtag.name = "title"
    subtagurl = str(subtag.get('href'))
    del subtag['href']
    del subtag['rel']
    del subtag['target']
    xmlFile.write("<item>\n" + str(subtag) + "\n")
    xmlFile.write("<link>" + str(subtagurl) + "</link>\n")
    pubdate = tag.find_previous('span')
    pubdate.name = "pubDate"
    #descpubdate[0].name = "description"
    pubdesc = tag.findAllNext('div', limit=4)
    pubdesc[3].name = "description"
    #TODO: check if pubdate contains no date but something like 'vandaag' or 'xx uur geleden'en dan uit een tag.[attribute] de juiste datum halen.
    xmlFile.write(str(pubdate) + "\n")
    xmlFile.write(str(pubdesc[3]) + "\n</item>\n")
    #print(tag.find_.('span'))

xmlFile.write("</channel>\n</rss>")
xmlFile.close()

from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup
#import os
from lxml import html
import time

#cwd = os.getcwd()
# print the current directory 
#print("Current working directory is:", cwd)

online = True
urlbase = 'https://advisories.ncsc.nl/advisory?id=NCSC-2021-0'
startkwetsnumb = 895
numbofpages = 8
totalkwets = startkwetsnumb + numbofpages - 1
#numbofpages includes the startpage

#Adjust filenames below before running this script!!
if online==False:
    htmlFile = open('oldreaderncscupdate20201104test.html')

# No editing needed below this line
# --------------------------------

baseName  = "ncscexport"
date = time.strftime('%Y%m%d')
fileName = baseName+date+".xml"
#fileName = baseName+date+"-0"+str(startkwetsnumb)+"-0"+str(totalkwets)+".xml"
xmlFile = open(fileName,'w')


# start of xml-content
xmlFile.write('<?xml version="1.0" encoding="UTF-8"?>\n\n<rss version="2.0">\n\n<channel>\n')

#scrape htmlpage from NCSC or import rss thru local xml-file
if online:
    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options=options)
    for page in range (numbofpages):
        urllink = urlbase + str(startkwetsnumb)
        try:
            #from url, enable if online
            driver.get(urllink)
            time.sleep(3)
            page = driver.page_source
            #driver.quit()
        except:
            print ("Url {} not available/found!".format(urllink))
        #toParse =  BeautifulSoup(page, "html.parser")
        tree = html.fromstring(page)
        title = tree.xpath('//*[@id="advisory_content"]/h1/text()')
        kwets = tree.xpath('//*[@id="ncsc_adv_history"]/tbody/tr[3]/td[4]/text()')
        pubdate_str = tree.xpath ('//*[@id="ncsc_adv_history"]/tbody/tr[3]/td[1]/text()')
    
        kans_str = tree.xpath ('//*[@id="ncsc_adv_history"]/tbody/tr[3]/td[2]/div/text()')
        kanswaarde = kans_str[1].strip("\n\t")
        kans=kanswaarde[0]
    
        schade_str = tree.xpath ('//*[@id="ncsc_adv_history"]/tbody/tr[3]/td[3]/div/text()')
        schadewaarde = schade_str[1].strip("\n\t")
        schade=schadewaarde[0]
    
        description_str = tree.xpath ('//*[@id="ncsc_adv_history"]/tbody/tr[5]/td[1]/p/text()')
    
        xmlFile.write("<item>\n<title>" + kwets[0] +" [" + kans.upper() + "/" + schade.upper() + "] " + title[0] +  "</title>\n")
        xmlFile.write("<link>" + urllink + "</link>\n")
    
        xmlFile.write("<pubDate>" + pubdate_str[0] + " 00:00</pubDate>\n")
        xmlFile.write("<description>" + description_str[0] + "</description>\n</item>\n")
        
        startkwetsnumb = startkwetsnumb + 1
    driver.quit()

else:
    #from a file, enable if local feed
    try:
        toParse = BeautifulSoup(htmlFile, "html.parser")
    except:
        print ("HTML-file {} not found!".format(htmlFile))
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

#end the xml and the file
xmlFile.write("</channel>\n</rss>")
xmlFile.close()

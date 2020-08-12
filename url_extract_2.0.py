from bs4 import BeautifulSoup
from urllib.request import urlopen
import re
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
import urllib
from fake_useragent import UserAgent
import openpyxl
import http.client
import feedparser
import csv

def read_csv(path = './input_file/all_urls.csv'):
    with open(path, encoding='utf-8') as f:
        fobj = csv.DictReader(f)

        for row in fobj:
            yield row

getALLlinks = []
Links = read_csv()
for id, link_i in enumerate(Links):
    # print(link_i["\ufeffurls"])
    getALLlinks.append(link_i["\ufeffurls"])

wb = openpyxl.Workbook()
sheet = wb.active
# sheet.title = 'Output'

#Add titles in the first row of each column
sheet.cell(row=1, column=1).value='id'
sheet.cell(row=1, column=2).value='url'
sheet.cell(row=1, column=3).value='title'
sheet.cell(row=1, column=4).value='corrected url'
sheet.cell(row=1, column=5).value='source'

wb.save('./output_file/scraped_pr_links.xlsx')


excelcounterrow = 2

processedlink = 0
skippedlink = 0
id = []

for lop in range(len(getALLlinks)):
    ua = UserAgent()
    req = Request(getALLlinks[lop])
    req.add_header('User-Agent', ua.random)
    processedlink = processedlink + 1

    print("Processing link no." + str(processedlink) + ": " + getALLlinks[lop])

    wbopen = openpyxl.load_workbook('./output_file/scraped_pr_links.xlsx')
    sheetopen = wbopen.active

    try:
        html_page = urlopen(req).read()
    #except (TimeoutError, ConnectionResetError, Exception, URLError, HTTPError, socket.timeout, http.client.RemoteDisconnected,http.client.HTTPException) as e:

    except (TimeoutError, ConnectionResetError, Exception, URLError, HTTPError, http.client.RemoteDisconnected, http.client.HTTPException, ConnectionError) as e:
        skippedlink = skippedlink + 1
        #html_page = urlopen(req)
        print('Unable to open link no.' + str(processedlink) + ": " + getALLlinks[lop])
        #sheet2.cell(row=excel2counterrow, column=1).value = getALLlinks[lop]
        #excel2counterrow = excel2counterrow+1

    except:
        skippedlink = skippedlink + 1
        #html_page = urlopen(req)
        print('Unable to open link no.' + str(processedlink) + ": " + getALLlinks[lop])

    else:
        soup = BeautifulSoup(html_page, features="html.parser")
        LinkYMLIndex = getALLlinks[lop]

        flag = 0

        for link in soup.findAll('a', href=True):
            # skip useless links

            if link['href'] == '' or link['href'].startswith(
                    '#'):  # if link['href'] == '' or link['href'].startswith('#'):     if link['href'] == '':
                continue

            # initialize the link
            thisLink = {
                'url': link['href'],
                'title': link.string,
            }
            actuallink = getALLlinks[lop]
            slashcounter = 0
            indexslash = 0

            if not actuallink.endswith("/"):
                actuallink = actuallink + "/"

            while slashcounter < 3:
                if(actuallink[indexslash] == "/"):
                    slashcounter = slashcounter + 1
                indexslash = indexslash + 1

            PLink = actuallink[:indexslash - 1]

            ModLink = thisLink['url'].strip()
            NewLink = ''

            if ModLink.startswith("/"):
                NewLink = PLink + ModLink
                # words = NewLink.split("/")
                # NewLink = "/".join(sorted(set(words), key=words.index))
            else:
                NewLink = ModLink

            if thisLink['title'] is None:
                # check for text inside the link
                if len(link.contents):
                    thisLink['title'] = ' '.join(link.stripped_strings)
            if thisLink['title'] is None:
                # if there's *still* no title (empty tag), skip it
                continue
            # convert to something immutable for storage

            sheetopen.cell(row=excelcounterrow, column=2).value = thisLink['url'].strip().replace('','')
            sheetopen.cell(row=excelcounterrow, column=3).value = thisLink['title'].strip().replace('','')
            sheetopen.cell(row=excelcounterrow, column=4).value = NewLink.replace('','')
            sheetopen.cell(row=excelcounterrow, column=5).value = getALLlinks[lop]
            excelcounterrow = excelcounterrow+1
            flag = 1

        if flag == 0:
            d = feedparser.parse(LinkYMLIndex)
            for rss in d.entries:
               linkrss = rss.link
               linktitle = rss.title

               actuallink = getALLlinks[lop]
               slashcounter = 0
               indexslash = 0
               while slashcounter < 3:
                   if (actuallink[indexslash] == "/"):
                       slashcounter = slashcounter + 1
                   indexslash = indexslash + 1
               PLink = actuallink[:indexslash - 1]

               ModLink = linkrss
               NewLink = ''

               if ModLink.startswith("/"):
                   NewLink = PLink + ModLink
                   # words = NewLink.split("/")
                   # NewLink = "/".join(sorted(set(words), key=words.index))
               else:
                   NewLink = ModLink

               sheetopen.cell(row=excelcounterrow, column=2).value = linkrss
               sheetopen.cell(row=excelcounterrow, column=3).value = linktitle
               sheetopen.cell(row=excelcounterrow, column=4).value = NewLink
               sheetopen.cell(row=excelcounterrow, column=5).value = getALLlinks[lop]
               excelcounterrow = excelcounterrow + 1


            #hashableLink = (thisLink['url'].strip(),thisLink['title'].strip(),NewLink,getALLlinks[lop])
            # store the result
            #if hashableLink not in links:
                #links.append(hashableLink)

    for i in range(len(sheetopen['url']) - 1):
        i = i + 1
        sheetopen.cell(row=i + 1, column=1).value = i

    sheet.title = 'Output'
    wbopen.save('./output_file/scraped_pr_links.xlsx')

print("\nTotal Processed Links: " + str(processedlink))
print("Total Unprocessed Links: " + str(skippedlink))
print("Output File Generated")

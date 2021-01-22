import requests
import xlsxwriter
from bs4 import BeautifulSoup
from urllib.request import urlopen

URL = 'https://rpa.hybrydoweit.pl/#articles'
page2 = urlopen(URL)
page = requests.get(URL)

titles= []
sections=[]
links=[]

soup = BeautifulSoup(page.content, 'html.parser')

articles = soup.find_all("article", {"class": "rpa-article-card"})

for article in articles:
    articlecards = article.find("div", {"class": "rpa-article-card__caption"})
    sectioncard = articlecards.find("ul", {"class": "rpa-article-card__metadata"})
    if sectioncard is not None:
        titles.append(articlecards.find("h3", {"class": "rpa-article-card__title"}).text)
        links.append(article.find('a')['href'])
        sections.append(sectioncard.find("li", {"class": "rpa-article-card__metadata-item"}).text.split(":",1)[1])

row = 0
col = 0
workbook = xlsxwriter.Workbook('articles.xlsx')
worksheet = workbook.add_worksheet('Articles')

worksheet.write(row, col, "Tytuł")
worksheet.write(row, col+1, "Branża/dział")
worksheet.write(row, col+2, "Link do artykułu")

listlength = len(titles)
print(listlength)
for i in range(listlength):
    row += 1
    worksheet.write(row, col, titles[i])
    worksheet.write(row, col+1, sections[i])
    worksheet.write(row, col+2, links[i])

workbook.close()

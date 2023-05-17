from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import xlsxwriter
import requests


def GetInfo(URL, row, worksheet):
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, "html.parser")
    name = soup.find("h2", class_="article-header").text
    #annotation_div = soup.find("div", class_="annot")
    annotation = soup.find("div", class_="annot").text
    keywords_div = soup.find("div", class_="keywords")
    keywords = keywords_div.find_all("a")
    keywords_str = ""
    for keyword in keywords:
        keywords_str += keyword.text
        keywords_str += ", "
    worksheet.write(row, 0, name)
    worksheet.write(row, 1, annotation)
    worksheet.write(row, 2, keywords_str)


def ParseVestnik(URL, num, workbook):

    worksheet_name = "Vestnik" + str(num)
    worksheet = workbook.add_worksheet(name=worksheet_name)
    row = 0
    worksheet.write(row, 0, "Название")
    worksheet.write(row, 1, "Аннотация")
    worksheet.write(row, 2, "Ключевые слова")
    worksheet.write(row, 3, "Авторы")
    row += 1

    page = requests.get(URL)
    soup = BeautifulSoup(page.content, "html.parser")
    all_divs = soup.find_all("div", class_="link")
    links_info = []
    for div in all_divs:
        links = div.find_all("a", href=True)
        for link in links:
            if not "persons" in link['href']:
                all_i = div.find_all("i")
                authors = ""
                for i in all_i:
                    authors += i.text

                worksheet.write(row, 4, authors)
                row += 1
                links_info.append(link['href'])

    row = 1
    for link in links_info:
        GetInfo(link, row, worksheet)
        row += 1

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    workbook = xlsxwriter.Workbook('../Journals.xlsx')
    links = ["https://iorj.hse.ru/2022-17-1.html", "https://iorj.hse.ru/2022-17-2.html", "https://iorj.hse.ru/2022-17-3.html", "https://iorj.hse.ru/2022-17-4.html"]
    i = 0
    for link in links:
        ParseVestnik(link, i, workbook)
        i += 1
    workbook.close()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

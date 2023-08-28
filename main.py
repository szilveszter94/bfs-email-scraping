import openpyxl as openpyxl
import requests
from bs4 import BeautifulSoup
import math

FIRST_PAGE = 1 # modify if you don't want to start from the first page
LAST_PAGE = False # modify only if you don't want to load all pages, else you can provide how many pages you want
ITEMS_PER_PAGE = 25  # items / page
# modify this link and after the '&page=' put '{}' because it's a dynamic url
BASE_LINK = "https://www.zoznam.sk/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/A/sekcia.fcgi?sid=1173&so=&page={}&desc=&shops=&kraj=&okres=&cast=&attr="
BASE_FIRMA_LINK = "https://www.zoznam.sk"   # base link for profiles (do not modify)
EXCEL_FILE_NAME = "email_by_name_0_9.xlsx"  # here you must provide the Excel file name

ALL_FIRMA_LINKS = []
ALL_EMAILS = []

# calculate the last page if not provided
if not LAST_PAGE:
    pageNumber = requests.get(BASE_LINK.format(FIRST_PAGE))
    soup = BeautifulSoup(pageNumber.content, "html.parser")
    page_num = soup.find('small').text
    formatted_page_number = int(page_num.strip("()"))
    LAST_PAGE = math.ceil(formatted_page_number / ITEMS_PER_PAGE)

# load the main pages and save the firma links into an array
for page_number in range(FIRST_PAGE, LAST_PAGE + 1):
    print(f'load pages: {round(page_number * 100 / LAST_PAGE, 2)} %')
    page_url = BASE_LINK.format(page_number)
    response = requests.get(page_url)
    soup = BeautifulSoup(response.content, "html.parser")
    elements = soup.find_all("a", class_="link_title")
    for i in elements:
        ALL_FIRMA_LINKS.append(f'{BASE_FIRMA_LINK}{i.get("href")}')

# loop through all firma links and extract the emails
link_counter = 0
ALL_PAGES = len(ALL_FIRMA_LINKS)
for link in ALL_FIRMA_LINKS:
    print(f'Load emails: {round(link_counter * 100 / ALL_PAGES, 2)} %')
    response = requests.get(link)
    soup = BeautifulSoup(response.content, "html.parser")
    mails = soup.select(".profile .row .col-sm-9 a")
    for i in mails:
        mail = i.text
        if "@" in mail:
            ALL_EMAILS.append([mail])
            break
    link_counter += 1

# save all emails into an Excel file
workbook = openpyxl.Workbook()
sheet = workbook.active
# write the Excel file
for row_data in ALL_EMAILS:
    sheet.append([str(cell) for cell in row_data])
workbook.save(EXCEL_FILE_NAME)


__author__ = 'kudinovdenis'
import xml.etree.ElementTree as ET
import os
import csv
import datetime

import re

TAG_RE = re.compile(r'<[^>]+>')

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

def remove_tags(text):
    return TAG_RE.sub('', text)


def unescape(s):
    s = s.replace("&lt;", "<")
    s = s.replace("&gt;", ">")
    # this has to be last:
    s = s.replace("&amp;", "&")
    return s


# 20140722T004122Z
# year = 2014
# month = 07
#day = 22
#hour = 00
#minute = 41
#second = 22
def parseTimestamp(timestamp):
    year = int(timestamp[:4])
    month = int(timestamp[4:-10])
    day = int(timestamp[6:-8])
    hour = int(timestamp[9:-5])
    minute = int(timestamp[11:-3])
    second = int(timestamp[13:-1])
    time = datetime.datetime(year=year, month=month, day=day, hour=hour, minute=minute, second=second)
    return time

# словарь с данными о книге
book_info_array = []
book_info = {}

# два счётчика -- всех книг и повторяющихся книг
all_books_counter = 0
double_books_counter = 0
#счётчик обновлённых книг
updated_books_counter = 0

# ISBN книги
ISBN = 0

#словарь для хранения данных о книгах: ISBN/словарь с данными
all_books_info = {}

#флаг того, что таймштамп в массиве для ISBN был найден(значит, такую книгу нужно не добавить, а проверить на таймштамп)
timestamp_found = False

# Часть 1 -- парсинг XML (открытие папки XML и всех её подпапок)
absolute_XML_Path = "/Users/kudinovdenis/Documents/other/onix-2-csv-Parser/XML"

for file in os.listdir(absolute_XML_Path):
    if not ".DS_Store" in file:
        for xml_file in os.listdir(os.path.join(absolute_XML_Path, file)):
            if not ".DS_Store" in xml_file:
                path = (os.path.join(file, xml_file))
                path = "XML/" + path
                tree = ET.parse(path)
                root = tree.getroot()
                #print(root)
                for lev0 in root:

                    # считывание таймштампа из <ONIXMessage>
                    if lev0.tag == "Header":
                        for lev1 in lev0:
                            if lev1.tag == "SentDateTime":
                                current_timestamp = lev0.find("SentDateTime").text
                                current_timestamp = parseTimestamp(current_timestamp)

                    #print(lev0.tag, ": ", root.find(lev0.tag).text)
                    book_info = {lev0.tag: root.find(lev0.tag).text}
                    book_info_array.append(book_info)
                    for lev1 in lev0:
                        #print("\t", lev1.tag, ": ", lev0.find(lev1.tag).text)
                        book_info = {lev1.tag: lev0.find(lev1.tag).text}
                        if lev1.tag == "RecordReference":
                            ISBN = lev0.find('RecordReference').text[:-5]
                        if not lev0.tag == "Header":
                            book_info_array.append(book_info)
                        for lev2 in lev1:
                            #print("\t\t", lev2.tag, ": ", lev1.find(lev2.tag).text)
                            book_info = {lev2.tag: lev1.find(lev2.tag).text}
                            book_info_array.append(book_info)
                            for lev3 in lev2:
                                #print("\t\t\t", lev3.tag, ": ", lev2.find(lev3.tag).text)
                                book_info = {lev3.tag: lev2.find(lev3.tag).text}
                                book_info_array.append(book_info)
                                for lev4 in lev3:
                                    #print("\t\t\t\t", lev4.tag, ": ", lev3.find(lev4.tag).text)
                                    book_info = {lev4.tag: lev3.find(lev4.tag).text}
                                    book_info_array.append(book_info)
                    #добавление только при условии, что встретилась новая книга


                    # таймштамп текущей книги
                    book_info = {"SentDateTime": current_timestamp}
                    book_info_array.append(book_info)

                    if lev0.tag == "Product":
                        if not ISBN in all_books_info.keys():
                            all_books_info[ISBN] = book_info_array
                        else:
                            double_books_counter += 1
                            # поиск таймштампа в массиве книг. Если там он меньше, чем в текущей книге, то удаляем элемент из массива книг
                            # и добавляем туда новый (текущий) элемент вместо него
                            for info in all_books_info[ISBN]:
                                if "SentDateTime" in info:
                                    in_array_timestamp = info["SentDateTime"]
                                    if in_array_timestamp < current_timestamp:
                                        all_books_info.__delitem__(ISBN)
                                        all_books_info[ISBN] = book_info_array
                                        updated_books_counter += 1
                                        break
                                    break
                        book_info_array = []
                        all_books_counter += 1

# Часть 2 -- парсинг csv
absolute_CSV_path = "/Users/kudinovdenis/Documents/other/onix-2-csv-Parser/CSV"

# берутся все csv файлы и в них пропускается первая строка, далее идут попытки взять все ISBN
# на выходе имеем массив ISBN книг, которые нужно удалить
# ISBN_to_delete -- массив с ISBN для удаления
ISBN_to_delete = []
for file in os.listdir(absolute_CSV_path):
    if not ".DS_Store" in file:
        path = "CSV/" + file
        with open(path) as csvfile:
            rows_counter = 0
            reader = csv.reader(csvfile)
            for row in reader:
                rows_counter += 1
                if rows_counter > 2 and len(row) > 0:
                    ISBN_to_delete.append(row[0])
#print(ISBN_to_delete)


# Часть 3 -- удаление книг, которые были в csv из словаря книг
before_count = len(all_books_info)
for key in ISBN_to_delete:
    all_books_info.__delitem__(key)

# Часть 4 -- создание csv
f = open('out.csv', 'w')

# заполнение названий полей в Excel
f.write("ISBN;Title;SubTitle;Regions;Prise USD;Currency;Text\n")

bookinfo = {}

#i -- счетчик "плохих" книг
i = 0

# parts -- массив с выходными колонками csv
parts = []
# important -- массив, при отсутствии в котором хотя бы одного элемента -- книги не попадуют в продажу
important = []
# во всех книгах выбираем одну по key(ISBN)
for key in all_books_info.keys():
    f.write(key)
    f.write(";")
    publisher = ''
    title = ''
    subtitle = ''
    region = ''
    priceAmount = ''
    currencyCode = ''
    text = ''
    #important!
    recordReference = ''
    notificationType = ''
    idValue = ''
    productFrom = ''
    titleText = ''
    bibliographicalNote = ''
    rightsCountry = ''
    priceEffectiveUntil = ''
    # book status
    publishingStatus = ''
    sentDateTime = ''

    # и берём инфу о ней
    # bookinfo здесь -- словарь из одной записи (но таких словарей много)
    for bookinfo in all_books_info[key]:

        # информация, которая будет выведена в Excel
        if "RecordSourceName" in bookinfo:
            publisher = bookinfo["RecordSourceName"]
        if "TitleText" in bookinfo:
            title = bookinfo["TitleText"]
        if "Subtitle" in bookinfo:
            subtitle = bookinfo["Subtitle"]
        if "RegionsIncluded" in bookinfo:
            region = bookinfo["RegionsIncluded"]
        if "PriceAmount" in bookinfo:
            priceAmount = bookinfo["PriceAmount"]
        if "CurrencyCode" in bookinfo:
            currencyCode = bookinfo["CurrencyCode"]
            priceAmount = priceAmount + currencyCode
        if "Text" in bookinfo:
            text = bookinfo["Text"]
            text = unescape(text)
            text = re.sub('<[^>]*>', '', text)
            text = re.sub('\n', '', text)
            text = re.sub(';', '', text)
            text = "\"" + text + "\""
            text = text.rstrip('\n')

        #important!
        if "RecordReference" in bookinfo:
            recordReference = bookinfo["RecordReference"]
        if "NotificationType" in bookinfo:
            notificationType = bookinfo["NotificationType"]
        if "IDValue" in bookinfo:
            idValue = bookinfo["IDValue"]
        if "ProductForm" in bookinfo:
            productFrom = bookinfo["ProductForm"]
        if "TitleText" in bookinfo:
            titleText = bookinfo["TitleText"]
        if "BibliographicalNote" in bookinfo:
            bibliographicalNote = bookinfo["BibliographicalNote"]
        if "RightsCountry" in bookinfo:
            rightsCountry = bookinfo["RightsCountry"]
        if "PriceEffectiveUntil" in bookinfo:
            priceEffectiveUntil = bookinfo["PriceEffectiveUntil"]

        if "PublishingStatus" in bookinfo:
            publishingStatus = bookinfo["PublishingStatus"]
        if "SentDateTime" in bookinfo:
            sentDateTime = bookinfo["SentDateTime"]

        """for bookinfo_key in bookinfo.keys():
            f.write(bookinfo_key)
            f.write(";")
            try:
                f.write(bookinfo[bookinfo_key])
                f.write(";")
            except:
                print("Error")"""

    parts = []
    #parts.append(publisher)
    parts.append(title)
    parts.append(subtitle)
    parts.append(region)
    parts.append(priceAmount)
    parts.append(currencyCode)
    parts.append(text)
    #parts.append(publishingStatus)
    #parts.append(sentDateTime)

    important = []
    important.append(recordReference)
    important.append(notificationType)
    important.append(idValue)
    important.append(productFrom)
    important.append(titleText)
    important.append(bibliographicalNote)
    important.append(rightsCountry)
    important.append(priceEffectiveUntil)

    #проверка у книги наличия всех необходимых полей
    for important_field in important:
        if not important_field:
            #print("В описании книги отсутствуют необходимые поля!\n Книга с ISBN: ", key)
            i += 1
            break

    for part in parts:
        if part:
            f.write(part)
            f.write(";")
        else:
            f.write("n/a;")
    f.write("\n")
print("Количество плохих книг: ", i, " (книги, которые не удовлетворяют формату ONIX) ", bcolors.WARNING, "не критично", bcolors.ENDC)
print("Количество книг в массиве после обработки: ", len(all_books_info), " (книги, которые пришлось удалить из-за наличия их в CSV об удалении)")
print("Всего было книг: ", all_books_counter, "(вместе с дубликатами)")
print("Из них дубликатов: ", double_books_counter)
print("\n")
print(bcolors.OKGREEN, "Всего книг до обрабтки: ", before_count, " (реальное число книг)", bcolors.ENDC)
print(bcolors.OKBLUE, "Книги были обновлены: ", updated_books_counter, " раз", bcolors.ENDC)
print(bcolors.FAIL, "Книг было удалено из-за снятия с продажи (CSV): ", before_count - len(all_books_info), bcolors.ENDC)
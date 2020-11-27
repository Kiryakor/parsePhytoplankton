from docx2python import docx2python
from enum import Enum
import csv

class ParseEnum(Enum):
    """Перечисление для отслеживания таблицы, которая парсится в данный момент"""
    header = 1
    body = 2
    bottom = 3
    notGood = 4

class Header:
    """Класс для хранения свойств первой таблицы"""
    def __init__(self, water="", date="", station="", depth="", temperature="", alpha="", author=""):
        self.water = water
        self.date = date
        self.station = station
        self.depth = depth #глубина
        self.temperature = temperature
        self.alpha = alpha
        self.author = author

class Takson:
    """Класс для хранения свойств 2 и 3 таблицы"""
    def __init__(self, takson, counter, bioMassa, percentCounter, percentBioMassa):
        self.takson = takson
        self.counter = counter
        self.bioMassa = bioMassa
        self.percentCounter = percentCounter
        self.percentBioMassa = percentBioMassa

class Content:
    """Класс для ханения данных одной полной таблицы"""
    def __init__(self, head, body, bottom):
        self.head = head
        self.body = body
        self.bottom = bottom

class ParseSecondBlock:

    def __init__(self, path):
        self.path = path
        self.parseData = []
        self.parseState = ParseEnum.notGood
        self.doc_result = docx2python(path)

        self.startParse()

    def parse1(self, content, header):
        """Метод для парсинг первой таблицы"""
        if "Водоем:" in content:
            a = content.find("Дата")
            b = content.find("Станция")
            header.water = content[7:a].strip()
            header.date = content[a + 5:b].strip()
            header.station = content[b + 8:].strip()
        elif "Глубина:" in content:
            a = content.find("Температура")
            b = content.find("Прозрачность")
            header.depth = content[8:a].strip()
            header.temperature = content[a + 12:b].strip()
            header.alpha = content[b + 13:].strip()
        elif "Исполнитель:" in content:
            header.author = content[12:].strip()

    def parse3(self, content):
        """Метод для парсинг 2 и 3 таблицы"""
        if content[0][0] != "Отдел" and content[0][0] != "Таксон":
            return Takson(content[0][0], content[1][0], content[2][0], content[3][0], content[4][0])
        return 0

    def saveData(self):
        with open('header.tsv', 'wt') as out_file:
            tsv_writer = csv.writer(out_file, delimiter='\t')
            counter = 0
            for i in self.parseData:
                tsv_writer.writerow(
                    [counter, i.head.water, i.head.date, i.head.station, i.head.depth, i.head.temperature, i.head.alpha,
                     i.head.author])
                counter += 1

        with open('body.tsv', 'wt') as out_file:
            tsv_writer = csv.writer(out_file, delimiter='\t')
            counter = 0
            for i in self.parseData:
                for j in i.body:
                    tsv_writer.writerow([counter, j.takson, j.counter, j.bioMassa, j.percentCounter, j.percentBioMassa])
                counter += 1

        with open('bottom.tsv', 'wt') as out_file:
            tsv_writer = csv.writer(out_file, delimiter='\t')
            counter = 0
            for i in self.parseData:
                for j in i.bottom:
                    tsv_writer.writerow([counter, j.takson, j.counter, j.bioMassa, j.percentCounter, j.percentBioMassa])
                counter += 1

    def startParse(self):
        for j in self.doc_result.body:
            header = Header()
            bottom = []
            body = []

            for i in j:
                if i[0][0] == "":
                    continue

                if "Водоем:" in i[0][0]:
                    self.parseState = ParseEnum.header
                elif "Таксон" == i[0][0]:
                    self.parseState = ParseEnum.body
                elif "Отдел" == i[0][0]:
                    self.parseState = ParseEnum.bottom
                elif "Всего" == i[0][0]:
                    takson = self.parse3(i)
                    if takson != 0:
                        bottom.append(takson)
                    self.parseState = ParseEnum.notGood

                if self.parseState == ParseEnum.header:
                    self.parse1(i[0][0], header)
                elif self.parseState == ParseEnum.body or self.parseState == ParseEnum.bottom:
                    takson = self.parse3(i)
                    if takson != 0 and self.parseState == ParseEnum.body:
                        body.append(takson)
                    elif takson != 0 and self.parseState == ParseEnum.bottom:
                        bottom.append(takson)
                elif self.parseState == ParseEnum.notGood:
                    pass
                else:
                    print("Error in parseState")

            if header.author != "":
                self.parseData.append(Content(header, body, bottom))

            self.saveData()

ParseSecondBlock('mt.docx')

#!/usr/bin/python
# -*- coding: utf-8 -*-

import random
import datetime

import xlwt


class RandomManager(object):

    @staticmethod
    def getCellulare():
        return int('393%09d' % random.randint(0, 999999999))

    @staticmethod
    def getNome():
        c = 'bcdfghjklmnpqrstvwxyz'
        v = 'aeiou'
        ret = ''
        for _ in range(random.randint(2, 5)):
            ret += random.choice(c) + random.choice(v)
        return ret.title()

    @staticmethod
    def getSesso():
        return random.choice(['M', 'F'])

    @staticmethod
    def getIndirizzo():
        v = random.choice(['Via', 'Piazza', 'Viale'])
        n = random.randint(1, 100)
        return '%s %s, %d' % (v, RandomManager.getNome(), n)

    @staticmethod
    def getCitta():
        return random.choice(['Roma', 'Milano', 'Napoli'])

    @staticmethod
    def getProvincia():
        l = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        return random.choice(l) + random.choice(l)

    @staticmethod
    def getDataNascita():
        return datetime.datetime.now()

    @staticmethod
    def getTelefono():
        return '0%09d' % random.randint(0, 999999999)

    @staticmethod
    def getEmail():
        return '%s@%s.it' % (RandomManager.getNome().lower(),
                             RandomManager.getNome().lower())

    @staticmethod
    def getNote():
        ret = [RandomManager.getNome()]
        for _ in range(random.randint(0, 9)):
            ret.append(RandomManager.getNome().lower())
        return ' '.join(ret)

    @staticmethod
    def getTestaOCroce():
        return random.choice([True, False])


class RandomItem(dict):

    def __init__(self, *args, **kw):
        dict.__init__(self, *args, **kw)
        self.setAttribute('Cellulare', RandomManager.getCellulare(),
                          random=False)
        self.setAttribute('Nome', RandomManager.getNome())
        self.setAttribute('Cognome', RandomManager.getNome())
        self.setAttribute('Sesso', RandomManager.getSesso())
        self.setAttribute('Indirizzo', RandomManager.getIndirizzo())
        self.setAttribute('Citta', RandomManager.getCitta())
        self.setAttribute('Provincia', RandomManager.getProvincia())
        self.setAttribute('DataNascita', RandomManager.getDataNascita())
        self.setAttribute('Telefono', RandomManager.getTelefono())
        self.setAttribute('Email', RandomManager.getEmail())
        self.setAttribute('Note', RandomManager.getNote())
        self.setAttribute('SMScompleanno', 'X', depends='DataNascita')

    def setAttribute(
        self,
        attr,
        value,
        random=True,
        depends=None,
        ):

        if random and RandomManager.getTestaOCroce():
            return
        if depends is None or self.has_key(depends):
            self[attr] = value


styleD = xlwt.easyxf(num_format_str='dd/mm/yyyy')
styleH = xlwt.easyxf('font: name Arial, bold on')

header = (
    'Nome',
    'Cognome',
    'Cellulare',
    'Sesso',
    'Indirizzo',
    'Citta',
    'Provincia',
    'DataNascita',
    'Telefono',
    'Email',
    'Note',
    'SMScompleanno',
    )


class ExcelGenerator(object):

    def __init__(self):
        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet('Rubrica')
        self.currentRow = 0
        self.createHeader()

    def createHeader(self):
        for (i, h) in enumerate(header):
            self.ws.write(self.currentRow, i, h, styleH)
        self.currentRow += 1

    def insertItem(self, d):
        for (k, v) in d.iteritems():
            if isinstance(v, datetime.datetime):
                self.ws.write(self.currentRow, header.index(k), v,
                              styleD)
            else:
                self.ws.write(self.currentRow, header.index(k), v)
        self.currentRow += 1

    def save(self, fName):
        self.wb.save(fName)


if __name__ == '__main__':
    eg = ExcelGenerator()
    for _ in range(50000):
        eg.insertItem(RandomItem())
    eg.save('rubrica.xls')

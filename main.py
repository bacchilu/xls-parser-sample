#!/usr/bin/python
# -*- coding: utf-8 -*-

import datetime
import time

import xlrd


class PHPSerialize(object):

    def __init__(self, ast, coding='utf-8'):
        self.ast = list(ast)
        self.coding = coding
        self.header = self.ast[0].keys()

    def getHeader(self):
        a = ((i + 1, len(v), v) for (i, v) in enumerate(self.header))
        b = ''.join('i:%d;s:%d:"%s";' % t for t in a)
        return 'i:1;a:%d:{%s}' % (len(self.header), b)

    def encodeValue(self, v):
        if isinstance(v, unicode):
            return v.encode(self.coding)
        elif isinstance(v, datetime.datetime):
            return v.strftime('%d/%m/%Y')

    def getSingleRecord(self, item):
        a = ((k, v) for (k, v) in item.iteritems() if v is not None)

        b = ((self.header.index(k) + 1, len(self.encodeValue(v)),
              self.encodeValue(v)) for (k, v) in a)

        return ''.join('i:%d;s:%d:"%s";' % e for e in b)

    def getRecordCardinality(self, item):
        return len([e for e in item.values() if e is not None])

    def getContent(self):

        a = ((i + 2, self.getRecordCardinality(item),
              self.getSingleRecord(item)) for (i, item) in
            enumerate(self.ast))

        return ''.join('i:%d;a:%d:{%s}' % e for e in a)

    def dump(self):
        numRows = 's:7:"numRows";i:%d;' % len(self.ast)

        t = (len(self.ast) + 1, self.getHeader(), self.getContent())
        cells = 's:5:"cells";a:%d:{%s%s}' % t

        return 'a:2:{%s%s}' % (numRows, cells)


class JSONfy(object):

    def __init__(self, ast, coding='utf-8'):
        self.ast = ast
        self.coding = coding

    def dumpField(self, k, v):
        if v is None:
            v2 = u'null'.encode(self.coding)
        elif isinstance(v, unicode):
            v2 = '"%s"' % v.encode(self.coding)
        elif isinstance(v, datetime.datetime):
            v2 = str(int(time.mktime(v.timetuple()))).decode('utf-8')
        return '"%s": %s' % (k.encode(self.coding), v2)

    def dumpRecord(self, item):
        elems = (self.dumpField(k, v) for (k, v) in item.items())
        return '{%s}' % ', '.join(elems)

    def dump(self):
        c = ', '.join(self.dumpRecord(item) for item in self.ast)
        return '[%s]' % c


class ExcelParser(object):

    def __init__(self, fName):
        self.fName = fName

    def getCellData(
        self,
        sheet,
        x,
        y,
        ):

        t = sheet.cell_type(x, y)
        c = sheet.cell_value(x, y)
        if t == xlrd.XL_CELL_EMPTY:
            return None
        elif t == xlrd.XL_CELL_TEXT:
            return c
        elif t == xlrd.XL_CELL_NUMBER:
            return str(int(c)).decode('utf-8')
        elif t == xlrd.XL_CELL_DATE:
            tup = xlrd.xldate_as_tuple(c, sheet.book.datemode)
            return datetime.datetime(*tup)
        raise 'Errore'

    def parse(self):
        book = xlrd.open_workbook(self.fName)
        sheet = book.sheets()[0]

        header = [sheet.cell_value(0, col) for col in
                  range(sheet.ncols)]

        for row in range(sheet.nrows - 1):

            r = (self.getCellData(sheet, row + 1, col) for col in
                range(sheet.ncols))

            yield dict(zip(header, r))


if __name__ == '__main__':
    print PHPSerialize(ExcelParser('file_esempio.xls').parse()).dump()

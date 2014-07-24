#!/usr/bin/env python
# encoding: utf-8

import unittest
import reader


class ExcelProducer:

    def __init__(self):
        self.sheetnameList = []
        self.workbook = {}

    def AddSheet(self, sheetname):
        self.sheetnameList.append(sheetname)
        self.workbook[sheetname] = [[]]

    def AddCell(self, row, col, cell, sheetname):
        sheet = self._GetSheet(sheetname)

        self._InfillSheet(row, col, sheet)

        sheet[row][col] = cell

    def AddCells(self, cells, sheetname):
        for row in range(len(cells)):
            for col in range(len(cells[row])):
                self.AddCell(row, col, cells[row][col], sheetname)

    def GetCell(self, row, col, sheetname):
        sheet = self._GetSheet(sheetname)

        return sheet[row][col]

    def GetBook(self):
        book = []
        for sheetname in self.sheetnameList:
            sheet = self._GetSheet(sheetname)
            self._Infilling(sheet)
            book.append([sheetname, sheet])

        return book

    def GetSheetName(self):
        return self.sheetnameList

    def GetARow(self, row, sheetname):
        sheet = self._GetSheet(sheetname)

        self._Infilling(sheet)

        return sheet[row]

    def SetBook(self, book):
        for sheet in book:
            sheetname = sheet[0]
            cells = sheet[1]
            self.sheetnameList.append(sheetname)
            self.workbook[sheetname] = cells and cells or [[]]

    def InsertRow(self, row, aRow, sheetname):
        sheet = self._GetSheet(sheetname)
        self._InfillSheet(row, len(aRow), sheet) # infill rows
        sheet[:] = sheet[:row] + [aRow] + sheet[row:]

    def _InfillSheet(self, row, col, sheet):
        while len(sheet) < row + 1: sheet.append([])
        for aRow in sheet:
            while len(aRow) < col + 1: aRow.append('')

    def _Infilling(self, sheet):
        row = len(sheet) - 1
        col = max(map(len, sheet)) - 1
        self._InfillSheet(row, col, sheet)

    def _GetSheet(self, sheetname):
        sheet = self.workbook.get(sheetname)
        if sheet == None:
            raise ValueError("no {} sheet.".format(sheet))
        else:
            return sheet


class ProducerTest(unittest.TestCase):

    def setUp(self):
        self.excelProducer = ExcelProducer()
        self.excelProducer1 = ExcelProducer()
        xlsFile = reader.ExcelReader('../test/test.xls')
        self.excelProducer1.SetBook(xlsFile.Reads())

    def tearDown(self):
        pass

    def test_ASheet(self):
        sheetname = 'WTF'
        self.excelProducer.AddSheet(sheetname)
        self.assertEqual(self.excelProducer.GetBook(), [['WTF', [[]]]])

    def test_ManySheets(self):
        sheetnameList = ['W', 'T', 'F']
        map(lambda n: self.excelProducer.AddSheet(n), sheetnameList)
        self.assertEqual(self.excelProducer.GetBook(),
                [['W', [[]]],
                ['T', [[]]],
                ['F', [[]]]])

    def test_ACell(self):
        pos = (5, 5)
        cell = 'Apple'
        sheetname='WTF'
        self.excelProducer.AddSheet(sheetname)
        self.excelProducer.AddCell(*pos, cell=cell, sheetname=sheetname)
        self.assertEqual(self.excelProducer.GetCell(*pos, sheetname=sheetname), cell)

    def test_ManyCell(self):
        sheetname = 'W'
        self.excelProducer.AddSheet(sheetname)
        self.excelProducer.AddCell(1, 1, 'Apple', sheetname)
        self.excelProducer.AddCell(3, 6, 'Car', sheetname)
        self.excelProducer.AddCell(5, 4, 'Cake', sheetname)
        self.excelProducer.AddCell(0, 0, 'OK', sheetname)
        self.assertEqual(self.excelProducer.GetCell(1, 1, sheetname), 'Apple')
        self.assertEqual(self.excelProducer.GetCell(3, 6, sheetname), 'Car')
        self.assertEqual(self.excelProducer.GetCell(5, 4, sheetname), 'Cake')
        self.assertEqual(self.excelProducer.GetCell(0, 0, sheetname), 'OK')

    def test_GetBook(self):
        sheetname = 'Hello'
        cells = [['a', 'b'], [], ['', '', 'dog', '']]
        self.excelProducer.AddSheet(sheetname)
        for r in range(len(cells)):
            for c in range(len(cells[r])):
                self.excelProducer.AddCell(r, c, cells[r][c], sheetname)

        self.assertEqual(self.excelProducer.GetBook(), [[sheetname, [['a', 'b', '', ''],
                                                                     ['', '', '', ''],
                                                                     ['', '', 'dog', '']]]])

    def test_AddCells(self):
        sheetname = 'Yooo'
        cells = [['a', 'b', '', ''],
                ['', '', '', ''],
                ['', '', 'dog', '']]
        self.excelProducer.AddSheet(sheetname)
        self.excelProducer.AddCells(cells, sheetname)
        self.assertEqual(self.excelProducer.GetBook(), [[sheetname, cells]])

    def test_SetBook(self):
        xlsFIle = reader.ExcelReader('../test/test.xls')
        self.assertEqual(self.excelProducer1.GetBook(), xlsFIle.Reads())

    def test_GetSheet(self):
        self.assertEqual(self.excelProducer1.GetSheetName(), ['Sheet one'])

    def test_GetFristLine(self):
        sheetname = self.excelProducer1.GetSheetName()[0]
        self.assertEqual(self.excelProducer1.GetARow(0, sheetname), ['A1', 'B1'])
        self.assertNotEqual(self.excelProducer1.GetARow(1, sheetname), ['A1', 'B1'])
        self.excelProducer.AddSheet('Yooo')
        self.excelProducer.AddCells([['a', 'b'], [], ['', '', 'c']], 'Yooo')
        self.assertEqual(self.excelProducer.GetARow(0, 'Yooo'), ['a', 'b', ''])

    def test_InsertRow(self):
        row = 3
        aRow = ['a', '', 'b']
        sheetname = self.excelProducer1.GetSheetName()[0]
        self.excelProducer1.InsertRow(row, aRow, sheetname)
        self.assertEqual(self.excelProducer1.GetARow(row, sheetname), aRow)

if __name__ == '__main__':
    unittest.main()


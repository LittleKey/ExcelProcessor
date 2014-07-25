#!/usr/bin/env python
# encoding: utf-8

import xlrd
import unittest
import xlutils


class ExcelReader(object):

    def __init__(self, filename, formatting=False):
        self.formatting_info = formatting
        try:
            self.workbook = xlrd.open_workbook(filename, formatting_info=self.formatting_info)
        except NotImplementedError as e:
            print("[NotImplementedError]: {}".format(e.message))
            self.workbook = xlrd.open_workbook(filename)

    def Reads(self):
        return self._GetBook()

    def _GetBook(self):
        book = []
        for sheet in self._GetSheet():
            book.append(self._GetASheet(sheet))

        return book

    def _GetARow(self, row, sheet):
        aLine = []
        for col in range(sheet.ncols):
            aLine.append(self._GetACell(row, col, sheet))

        return aLine

    def _GetASheet(self, sheet):
        allCell = []
        for row in range(sheet.nrows):
            allCell.append(self._GetARow(row, sheet))

        allCell = allCell and allCell or [[]]
        return [sheet.name, allCell]

    def _GetSheet(self):
        for s in range(self.workbook.nsheets):
            yield self.workbook.sheet_by_index(s)

    def _GetACell(self, row, col, sheet):
        if self.formatting_info:
            return sheet.cell(row, col)
        else:
            return sheet.cell(row, col).value


class ReaderTest(unittest.TestCase):

    def setUp(self):
        self.excelFile1 = ExcelReader('../test/test.xls')
        self.excelFile2 = ExcelReader('../test/test_two.xlsx')
        self.excelFile3 = ExcelReader('../test/test.xls', formatting=True)

    def tearDown(self):
        pass

    def test_Read_xlsx(self):
        self.assertEqual(self.excelFile2.Reads(), [['Sheet1', [['Hey!']]]])

    def test_GetARow(self):
        self.assertEqual(self.excelFile1._GetARow(0, self.excelFile1._GetSheet().next()), ['A1', 'B1'])

    def test_GetBook(self):
        self.assertEqual(self.excelFile1.Reads(), [['Sheet one', [['A1', 'B1'], ['A2', 'B2']]],
                                                    ['Sheet2', [[]]],
                                                    ['Sheet3', [[]]]])

    def test_GetASheet(self):
        self.assertEqual(self.excelFile1._GetASheet(self.excelFile2._GetSheet().next()), ['Sheet1', [['Hey!']]])

    def test_GetBook_format(self):
        book = []
        for sheet in self.excelFile3.Reads():
            sheetname = sheet[0]
            cells = sheet[1]
            newCells = []
            for row in cells:
                newCells.append(map(lambda c: c.value, row))
            book.append([sheetname, newCells])

        self.assertEqual(book, self.excelFile1.Reads())

if __name__ == '__main__':
    unittest.main()


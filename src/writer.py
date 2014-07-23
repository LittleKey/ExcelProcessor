#!/usr/bin/env python
# encoding: utf-8

import xlwt
import unittest
import reader
import os
import producer


class ExcelWriter:

    def __init__(self, filename):
        self.workbook = xlwt.Workbook('utf-8')
        self.filename = filename

    def Write(self, book):
        for sheet in book:
            sheetname = sheet[0]
            allCells = sheet[1]
            self.WriteASheet(sheetname, allCells)

        self.Save()

    def WriteASheet(self, sheetname, allCells):
        newSheet = self.workbook.add_sheet(sheetname)
        self.WriteCells(newSheet, allCells, 0, 0)

        self.Save()

    def WriteCells(self, newSheet, allCells, r=0, c=0):
        pos = [r, c]
        for aRow in allCells:
            for cell in aRow:
                newSheet.write(*pos, label=cell)
                pos[1] += 1
            pos[0] += 1
            pos[1] = 0
        newSheet.flush_row_data()

        self.Save()

    def Save(self):
        self.workbook.save(self.filename)


class WriterTest(unittest.TestCase):

    def setUp(self):
        self.filename = '../test/test_tmp.xls'
        self.xlsFile = ExcelWriter(self.filename)
        pass

    def tearDown(self):
        try:
            os.remove(self.filename)
        except OSError as e:
            print(e.message)

    def test_Write_xls(self):
        #bookContext = [['Hello', [['1,A', '1,B'], ['2,A', '2,B']]]]
        bookContext = producer.ExcelProducer()
        bookContext.AddSheet('Hello')
        bookContext.AddCell(0, 0, '1,A', 'Hello')
        bookContext.AddCell(0, 1, '1,B', 'Hello')
        bookContext.AddCell(1, 3, '2,A', 'Hello')
        bookContext.AddCell(1, 1, '2,B', 'Hello')
        self.xlsFile.Write(bookContext.GetBook())
        self.assertXlsFileEqual(self.filename, bookContext.GetBook())

    def ftest_WriteASheet(self):
        sheetContext = ['Hello', [['1,A', '1,B'], ['2,A', '2,B']]]
        self.xlsFile.WriteASheet(sheetContext[0], sheetContext[1])
        self.assertXlsFileEqual(self.filename, [sheetContext])

    def ftest_WriteCells(self):
        cellsContext = [['2,A', '2,B']]
        sheet = self.xlsFile.workbook.add_sheet('Hello')
        self.xlsFile.WriteCells(sheet, cellsContext, r=1, c=0)
        self.assertXlsFileEqual(self.filename, [['Hello', [['', '']] + cellsContext]])

    def assertXlsFileEqual(self, filename, rvalue):
        self.assertEqual(reader.ExcelReader(filename).Reads(), rvalue)

if __name__ == '__main__':
    unittest.main()

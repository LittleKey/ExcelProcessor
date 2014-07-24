#!/usr/bin/env python
# encoding: utf-8

from reader import ExcelReader
from writer import ExcelWriter
from producer import ExcelProducer
import unittest
import os


class ExcelProcessor:

    def __init__(self, excelProducter):
        self.excel = excelProducter

    def Copy(self, oldRow, newRow, sheetname):
        aRow = self.excel.GetARow(oldRow, sheetname)
        self.excel.InsertRow(newRow, aRow, sheetname)

    def Save(self, filename):
        outFile = ExcelWriter(filename)
        outFile.Write(self.excel.GetBook())
        outFile.Save()


class ProcessorTest(unittest.TestCase):

    def setUp(self):
        self.excelFile = ExcelProducer()
        self.excelFile.SetBook(ExcelReader("../test/test.xls").Reads())
        self.excelProcessor = ExcelProcessor(self.excelFile)
        self.xlsFilename = "../test/tmp.xls"

    def tearDown(self):
        try:
            os.remove(self.xlsFilename)
        except OSError as e:
            print(e.message)

    def test_Copy(self):
        oriRow = 0
        newRow = 3
        sheetname = self.excelFile.GetSheetName()[0]
        self.excelProcessor.Copy(oriRow, newRow, sheetname)

        self.assertEqual(self.excelFile.GetARow(oriRow, sheetname), \
                        self.excelFile.GetARow(newRow, sheetname))

    def test_Save(self):
        sheetname = self.excelFile.GetSheetName()[0]
        self.excelProcessor.Copy(0, 3, sheetname)
        self.excelProcessor.Save(self.xlsFilename)
        self.assertEqual(self.excelFile.GetBook(), ExcelReader(self.xlsFilename).Reads())

if __name__ == '__main__':
    unittest.main()


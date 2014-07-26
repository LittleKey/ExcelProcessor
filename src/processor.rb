#!/usr/bin/env ruby
#encoding: utf-8

require 'spreadsheet'


module Excel
  class Processor
    def initialize
      @writeBook = Spreadsheet::Workbook.new
      @default_sheet = nil
    end

    def add_sheet(sheetName)
      @default_sheet = @writeBook.create_worksheet :name => sheetName
    end

    def push_row(aRow)
      @default_sheet.insert_row(@default_sheet.row_count)
      row = @default_sheet.last_row

      aRow.length.times.each do |num|
        row.push aRow[num]
        row.set_format(num, aRow.format(num))
      end
    end

    def copy_row_from(rb, copyRange)
      rSheet = rb.worksheet(0)

      copyRange.each do |num|
        push_row(rSheet.row(num))
      end
    end

    def Save(filename)
      @writeBook.write filename
    end
  end
end

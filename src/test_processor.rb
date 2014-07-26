#!/usr/bin/env ruby
#encoding: utf-8

require_relative 'processor'
require_relative 'book'
require 'minitest/autorun'


module Excel
  class TestProcessor < Minitest::Test
    def setup
      @workbook = Book.new '6.xls'
      @processor = Processor.new
    end

    def teardown
    end

    def test_add_sheet
      @processor.add_sheet 'Sheet One'
      @processor.Save '7.xls'
      wSheet = get_sheet_from_file '7.xls'

      assert_equal(wSheet.name, 'Sheet One')
    end

    def test_push_row
      rRow = @workbook.worksheet(0).row(0)

      @processor.add_sheet 'Sheet One'
      @processor.push_row rRow
      @processor.Save('7.xls')

      wSheet = get_sheet_from_file('7.xls')

      assert_equal(wSheet.row(0), rRow)
    end

    def test_copy_row_from
      @processor.add_sheet 'Sheet one'
      @processor.copy_row_from(@workbook, 0..1)
      @processor.Save('7.xls')
      wSheet = get_sheet_from_file('7.xls')
      rSheet = @workbook.worksheet(0)

      (0..1).each do |num|
        assert_equal(rSheet.row(num), wSheet.row(wSheet.last_row_index - 1 + num))
      end
    end

    def get_sheet_from_file(filename)
      wb = Book.new filename
      wb.worksheet 0
    end
  end
end

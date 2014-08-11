#!/usr/bin/env ruby
#encoding: utf-8

require 'minitest/autorun'
require_relative 'book'


module Excel
  class TestBook < Minitest::Test
    def setup
      @workbook = Book.new '../test/test_two.xlsx'
    end

    def teardown
    end

    def test_get_xlsx_sheet
      assert_equal(@workbook.worksheets[0].sheet_name, 'Sheet1')
    end

    def test_cell_value
      assert_equal(@workbook.worksheets[0].sheet_data[0][0].value, 'Hey!')
    end

    def test_excel_new
      Book.new 'No such file.xlsx'
    end

    def test_undefined_method
      assert_raises(NoMethodError) { @workbook.undefined_method() }
    end
  end
end

#!/usr/bin/env ruby
#encoding: utf-8

require 'minitest/autorun'
require_relative 'book'


module Excel
  class TestBook < Minitest::Test
    def setup
      @workbook = Book.new '6.xls'
    end

    def teardown
    end

    def test_get_sheet
      assert_equal(@workbook.worksheet(0).name, 'Sheet1')
      assert_equal(@workbook.worksheet(1).name, 'Sheet2')
      assert_equal(@workbook.worksheet(2).name, 'Sheet3')
    end

    def test_undefined_method
      assert_raises(NoMethodError) { @workbook.undefined_method() }
    end
  end
end

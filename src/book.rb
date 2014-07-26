#!/usr/bin/env ruby
#encoding: utf-8

#require 'test/unit'
require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'


module Excel
  class Book
    attr_reader :book
    def initialize(filename='')
      if File.exist? filename
        @book = Spreadsheet.open filename
      else
        @book = Spreadsheet::Workbook.new
      end
    end

    def method_missing(name, *args)
      if @book.methods.include? name
        @book.send(name, *args)
      else
        super
      end
    end
  end
end

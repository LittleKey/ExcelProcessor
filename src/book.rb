#!/usr/bin/env ruby
#encoding: utf-8

#require 'test/unit'
require 'spreadsheet'
require 'roo'
require 'rubyXL'

Spreadsheet.client_encoding = 'UTF-8'


module Excel
  class Book
    attr_reader :book
    def initialize(filename='')
      @filename = filename
      if File.exist? @filename
        @book = RubyXL::Parser.parse @filename
      else
        @book = RubyXL::Workbook.new
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

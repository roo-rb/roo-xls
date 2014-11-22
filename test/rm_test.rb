require 'spreadsheet'

book = Spreadsheet.open 'tmp.xls'
sheet = book.worksheet 0
sheet.each { |row| puts row[0] }

FileUtils.rm('tmp.xls')

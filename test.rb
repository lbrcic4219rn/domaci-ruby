require 'roo'
require 'roo-xls'
require './excelColumn'

class ExcelParser
  include Enumerable

  def initialize(path)
    @table = Roo::Spreadsheet.open(path, { :expand_merged_ranges => true })
    @columns = Hash.new
    @table.row(1).each_with_index { |header, i| @columns[header.downcase.gsub(' ', '_')] = i+1 }
  end

  def +(table2)
    if(@table.row(1).eql?(table2.table.row(1)))
      return  @table.parse + table2.table.parse
    end
    return nil
  end

  def -(table2)
    if(@table.row(1).eql?(table2.table.row(1)))
      return @table.parse - table2.table.parse
    end
    return nil
  end

  def method_missing(header)
    Col.new(@table.column(@columns[header.to_s]).drop(1), @table, @columns[header.to_s])
  end

  def [](key)
    #drop obrise prvi element niza
    return @table.column(@columns[key.to_s.downcase.gsub(' ','_')]).drop(1)
  end

  def row(num)
    return @table.row(num)
  end

  def each
    yield @table.parse
  end

  def table
    @table
  end

  def to_s
    @table.parse
  end
end


test_table = ExcelParser.new("./ruby_test.xlsx")
#test_table = ExcelParser.new("./testxls.xls")
test_table2 = ExcelParser.new("./test2.xlsx")

test_table.table.set(2, 1, "testSET")
# ispis table procitane iz excela
p test_table.to_s

#ispis headera tabele i ostalim redovima tabele preko indeksa
puts test_table.row(1)

#pristup kolonama tabele
p test_table["treca_kolona"]

#pristup elementima kolone
p test_table["treca_kolona"][2]

#pristup kolonama pomocu metoda
p test_table.prva_kolona.to_s

#suma kolone pomocu metoda
p test_table.prva_kolona.sum

#uzimanje reda pomocu istoimene funkcije iz kolone
p test_table.druga_kolona.bla

puts "rezultat sabiranje"
p test_table2+test_table

puts "rezultat oduzimanje"
p test_table2-test_table




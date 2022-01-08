require 'roo'

class Col

  def initialize(column, table, colindex)
    @column = column
    @table = table
    @colindex = colindex
  end

  def sum
    #sabiramo elemente niza
    @column.inject(0){ |sum, x| sum + x.to_f }
  end

  def to_s
    @column
  end

  def method_missing(m)
    @column.each_with_index do |col, i|
      if col.to_s.downcase.gsub(' ', '_').eql?(m.to_s.downcase.gsub(' ', '_'))
        #dodajemo 2 zato sto je i = 0 i preskacemo heder
        return @table.row(i+2)
      end
    end
  end

end
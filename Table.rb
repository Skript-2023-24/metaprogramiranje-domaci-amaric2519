# frozen_string_literal: true

class Table
  include Enumerable

  attr_accessor :rows, :cols, :izuzeci

  def initialize(worksheet)
    @table_matrix = []
    @rows = []
    @cols = []
    @izuzeci = %w[total subtotal]
    @worksheet = worksheet
    @worksheet.rows.each { |row| @table_matrix.push(row) unless row.any? { |i| @izuzeci.include? i.downcase } }
    create_rows
    create_cols
    create_header_functions
  end

  # kreira metode na osnovu header kolona excel fajla metode se zapisuju camelcase-om
  def create_header_functions
    @table_matrix[0].each do |el|
      tmp = el.to_s.split
      tmp[0] = tmp[0].downcase
      define_singleton_method(tmp.join) do
        index = @table_matrix[0].find_index(el)
        return @cols[index]
      end
    end
  end

  def create_rows
    @table_matrix.each { |row| @rows.push(Line.new(row, self, @rows.length)) }
  end

  def create_cols
    @table_matrix.transpose.each { |col| @cols.push(Line.new(col, self, @cols.length)) }
  end

  def cela_tabela
    @table_matrix
  end

  def row(par)
    @table_matrix[par]
  end

  def [](kol)
    index = @table_matrix[0].find_index(kol)
    @cols[index]
  end

  def refresh(arr, index)
    @rows.each_with_index do |el, id|
      el.change_value_at(index, arr[id])
      @worksheet[id+1, index+1] = arr[id]
      @worksheet.save
    end
    @table_matrix = []
    @rows.each { |el| @table_matrix.push(el.inspect) }
  end

  def method_missing(method_name)
    row = @table_matrix.detect{|aa| aa.include?(method_name.to_s)}
    @cols[row.index(method_name.to_s)]
  end
end

# Klasa row predstavlja jedan red u eksel fajlu nad kojim mogu da se vrse operacije
class Line
  include Enumerable

  def initialize(row, table, index)
    @index = index
    @table = table
    @arr = []
    row.any? { |w| w.count('a-zA-Z').positive? } ? row.each { |el| @arr.push(el) } : row.map { |el| @arr.push(el.to_i) }
  end

  def avg
    (self.sum / @arr.length).to_f.round(2)
  end

  def change_value_at(index, value)
    @arr[index] = value
  end

  def []=(brackets, data)
    @arr[brackets] = data
    @table.refresh(@arr, @index)
  end

  def [](brackets)
    @arr[brackets]
  end

  def each
    i = 0
    while i < @arr.length
      yield @arr[i]
      i += 1
    end
  end

  def get_at(index)
    @arr[index]
  end

  def find_index(word)
    @arr.find_index(word)
  end

  def check_if_contains(word)
    @arr.include? word
  end

  def method_missing(method_name)
    @table.row(@arr.find_index(method_name.to_s)) if @arr.include? method_name.to_s
  end

  def inspect
    @arr
  end

  def sum
    sum = 0
    @arr.each { |el| sum += el.to_i }
    sum
  end
end

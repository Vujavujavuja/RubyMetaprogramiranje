require 'google_drive'
require 'forwardable'
require 'ostruct'

#Sheet class
class Sheet
  include Enumerable

  def initialize(sheet, merged_cells = [])
    @sheet = sheet
    @merged_cells = merged_cells
    @headers = get_headers
    define_column_methods
  end


  def metged_cells
    @merged_cells
  end

  def headers
    @headers
  end

  def sheet
    @sheet
  end

  def rows
    @sheet.rows
  end

  def get_headers
    @sheet.rows.first.each_with_index.with_object({}) do |(cell, col_index), headers|
      headers[cell.to_s.strip.downcase] = col_index unless cell.to_s.strip.empty?
    end
  end

  #1) Biblioteka može da vrati dvodimenzioni niz sa vrednostima tabele
  def to_a
    @sheet.rows.drop(1).reject { |row| total_subtotal?(row) }
  end

  #2) Moguće je pristupati redu preko t.row(1), i pristup njegovim elementima po sintaksi niza.
  def row(row_index)
    @sheet.rows[row_index]
  end

  def get_column(header)
    col_index = @headers[header.to_s.strip.downcase]
    return [] unless col_index
    @sheet.rows.drop(1).map { |row| row[col_index] }
  end

  #3) Mora biti implementiran Enumerable modul(each funkcija), gde se vraćaju sve ćelije unutar tabele, sa leva na desno.
  def each
    @sheet.rows.drop(1).each_with_index do |row, row_index|
      row.each_with_index do |cell, col_index|
        if is_merged?(row_index + 1, col_index + 1) | total_subtotal?(row)
          next
        else
          yield cell
        end
      end
    end
  end


  #4) Biblioteka treba da vodi računa o merge-ovanim poljima
  def is_merged?(row_index, col_index)
    @merged_cells.any? do |merged|
      (merged.start_row..merged.end_row).include?(row_index) &&
          (merged.start_col..merged.end_col).include?(col_index)
    end
  end

  def[](column_header)
    Column.new(self, column_header, @headers[column_header.to_s.strip.downcase], @merged_cells)
  end

  def num_rows
    @sheet.num_rows
  end

  def insert_rows(row_index, rows)
    @sheet.insert_rows(row_index, rows)
    @sheet.save
  end

  #8) Sabiranje tabela
  def +(other_sheet)
    raise "Headers do not match" unless headers_match?(other_sheet)

    new_rows = @sheet.rows[1..].reject { |row| total_subtotal?(row) } + other_sheet.rows[1..].reject { |row| total_subtotal?(row) }
    new_sheet = Sheet.new(@sheet.dup)
    new_sheet.insert_rows(new_sheet.num_rows + 1, new_rows)
    new_sheet
  end

  #9) Oduzimanje tabela
  def -(other_sheet)
    raise "Headers do not match" unless headers_match?(other_sheet)

    new_rows = @sheet.rows[1..] - other_sheet.rows[1..]
    new_sheet = Sheet.new(@sheet.dup)
    new_sheet.insert_rows(new_sheet.num_rows + 1, new_rows)
    new_sheet
  end

  def rows_without_empty
    @sheet.rows.reject { |row| row_empty?(row) }
  end

  private

  def row_empty?(row)
    row.all? { |cell| cell.to_s.strip.empty? }
  end


  def headers_match?(other_sheet)
    self.headers == other_sheet.headers
  end


  #6) Pristup koloni preko istoimenih metoda
  def define_column_methods
    @headers.each_key do |header|
      method_name = header_to_method_name(header)
      define_singleton_method(method_name) do
        self[header]
      end
    end
  end

  def header_to_method_name(header)
    header.split('_').map(&:capitalize).join
  end

end



#4) Biblioteka treba da vodi računa o merge-ovanim poljima
class Merged
  attr_accessor :start_row, :start_col, :end_row, :end_col
  def initialize(start_row, start_col, end_row, end_col)
    @start_row = start_row
    @start_col = start_col
    @end_row = end_row
    @end_col = end_col
  end

end

class Column
  include Enumerable

  attr_accessor :header

  def initialize(sheet, header, column_index, merged_cells = [])
    @header = header
    @sheet = sheet
    @column_index = column_index
    @merged_cells = merged_cells
  end

  def sum
    extract_values.sum
  end

  def avg
    values = extract_values
    return 0 if values.empty?
    values.sum / values.size.to_f
  end

  def extract_values
    @sheet.rows.drop(1).map do |row|
      value = row[@column_index]
      allowed_value?(value) ? value.to_f : 0
    end
  end

  def allowed_value?(value)
    true if Float(value) rescue false
  end

  def [](row_index)
    cell_value(row_index + 1)
  end

  def []=(row_index, value)
    set_cell_value(row_index + 1, value)
  end

  def cell_value(row_index)
    @sheet.sheet[row_index, @column_index + 1]
  end

  def set_cell_value(row_index, value)
    @sheet.sheet[row_index, @column_index + 1] = value
    @sheet.sheet.save
  end

  def inspect
    values = @sheet.rows.drop(1).map { |row| row[@column_index] }.join(", ")
    "Header: #{@header}, Values: #{values}"
  end

  #6) Pristup koloni preko istoimenih metoda
  def method_missing(method_name, *arguments, &block)
    key = method_name.to_s.downcase
    row = @sheet.rows.find { |r| r[@column_index].to_s.downcase == key }
    row ? row : super
  end

  def respond_to_missing?(method_name, include_private = false)
    key = method_name.to_s.downcase
    @sheet.rows.any? { |r| r[@column_index].to_s.downcase == key } || super
  end

  def each
    @sheet.rows.drop(1).each do |row|
      next if total_subtotal?(row)
      yield row[@column_index]
    end
  end

end

#7) Prepoznavanje total i subtotal keyword-a
def total_subtotal?(row)
  row.any? { |cell| cell.to_s.downcase.match?(/\b(total|subtotal)\b/) }
end

def print_sheet(sheet)
  sheet.sheet.rows.each do |row|
    if row != sheet.sheet.rows.first
      row.each do |cell|
        print "|" + cell.to_s
      end
      puts "|"
    end
    end
end

def load_worksheet(session, key, index)
  sheet = session.spreadsheet_by_key(key).worksheets[index]
  raise "Worksheet not found for key: #{key}, index: #{index}" if sheet.nil?
  sheet
end

#MAIN?
session = GoogleDrive::Session.from_config("config.json")

sheet = load_worksheet(session, "15tGlSqtW0QdC4gUWRgOlsGOLa4CpCvs-PX6VjJjW2uw", 0)
merged_list = []
m1 = Merged.new(3, 4, 4, 4)
m2 = Merged.new(2, 4, 2, 4)
merged_list << m1
merged_list << m2
tabela = Sheet.new(sheet, merged_list)

sheet = load_worksheet(session, "1zYMBPZ1TML7Api8DHAgdE4rzcphWEHxV-Nbk_jHIxzE", 0)
tabela1 = Sheet.new(sheet, nil)

sheet = load_worksheet(session, "1A2yrG-R6lXU2FMzAWRfYr1mzxF2dKLPtRVF1XAAZJ-g", 0)
tabela2 = Sheet.new(sheet, nil)

#puts "Tabela 1:"
#puts tabela.get_headers.inspect
#print_sheet(tabela)

puts "\n1) Dvodimenzioni niz sa vrednostima iz tabele:"
puts tabela.to_a.inspect

puts "\n2) Pristup redu i pristup njegovim elementima po sintaksi niza:"
puts tabela.row(1).inspect

puts "\n3) Enumerable modul(each funkcija), gde se vraćaju sve ćelije unutar tabele, sa leva na desno:"
tabela.each do |cell|
  print(cell + "-")
end

puts "\n\n4) Vodi računa o merge-ovanim poljima:"
puts tabela.metged_cells.inspect

puts "\n5) Pristup koloni po headeru:"
puts tabela["ime"].inspect
print("Pre: " + tabela["ime"][1].inspect)
#tabela["ime"][1] = "Pera"
print(" Posle:" + tabela["ime"][1].inspect + "\n")

puts "\n6) Pristup koloni preko istoimenih metoda:"
print("6.i) Prva kolona sum: ")
puts tabela1.Prvakolona.sum
print("6.i) Druga kolona avg: ")
puts tabela1.Drugakolona.avg
print("6.ii) Red preko vrednosti u celiji:  ")
print(tabela.Index.rn1422)
puts "\n6.iii) Kolona podrzava enumerable:  "
puts tabela1.Prvakolona.map { |cell| cell.to_i + 1 }.inspect

puts "\n7) Prepoznavanje total i subtotal keyword-a: "
print_sheet(tabela2)
puts tabela2.to_a.inspect

puts("\n8) Sabiranje tabela: ")
puts("Tabela 1:")
puts tabela1.to_a.inspect
puts("Tabela 2:")
puts tabela2.to_a.inspect
puts("Tabela 1 + Tabela 2:")
puts (tabela1 + tabela2).to_a.inspect

puts("\n9) Oduzimanje tabela: ")
puts("Tabela 1:")
puts tabela1.to_a.inspect
puts("Tabela 2:")
puts tabela2.to_a.inspect
puts("Tabela 1 - Tabela 2:")
puts (tabela1 - tabela2).to_a.inspect
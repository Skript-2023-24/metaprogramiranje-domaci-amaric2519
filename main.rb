require './connect'
require './Table'

spreadsheet = connect('rubyDomaci1')

worksheet1 = spreadsheet.worksheets.first
# worksheet2 = spreadsheet.worksheets.each { | | }

t = Table.new(worksheet1)
# t2 = Table.new(worksheet2)
p "Ispis tabele"
t.cela_tabela.each { |row| p row }
p "Red 0"
p t.row(0)
p "t['Prva Kolona']"
p t['Prva Kolona']
p "t['Prva Kolona'][1]"
p t['Prva Kolona'][1]
p "t['Prva Kolona'][2]= 2556"
p t['Prva Kolona'][2] = 2556
p "t['Prva Kolona'][1]"
p t['Prva Kolona'][1]
p "t['Prva Kolona']"
p t['Prva Kolona']
p "t.prvaKolona"
p t.prvaKolona
p "t.prvaKolona.sum"
p t.prvaKolona.sum
p "t.prvaKolona.avg"
p t.prvaKolona.avg
p "t.prvaKolona.map{ |c| c = c.to_i + 1 }"
p t.prvaKolona.map{ |c| c = c.to_i + 1 }
p "t.indeks.rn3019"
p t.indeks.rn3019

# t2.cela_tabela

worksheet1.save

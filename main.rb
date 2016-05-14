=begin
require 'writeexcel'

# Create a new Excel Workbook
workbook = WriteExcel.new('ruby.xls')

# Add worksheet(s)
worksheet  = workbook.add_worksheet
worksheet2 = workbook.add_worksheet

# Add and define a format
format = workbook.add_format
format.set_bold
format.set_color('red')
format.set_align('right')

# write a formatted and unformatted string.
worksheet.write(1, 1, 'Hi Excel.', format)  # cell B2
worksheet.write(2, 1, 'Hi Excel.')          # cell B3

# write a number and formula using A1 notation
worksheet.write('B4', 3.14159)
worksheet.write('B5', '=SIN(B4/4)')

# write to file
workbook.close
=end

require 'Mdb'

database = Mdb.open('./main.mdb')

# list tables in the database
p database.tables 

# read the records in a table
#database[:Movies]
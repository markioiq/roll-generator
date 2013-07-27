# vim:set fileencoding=Windows-31J:

$LOAD_PATH << File.dirname(File.expand_path(__FILE__))

require 'win32ole'
require 'SeminarClass'

def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

filename = getAbsolutePath(File.dirname(File.expand_path(__FILE__)) + "/座席番号・定員.xlsx")

xl = WIN32OLE.new('Excel.Application')

book = xl.Workbooks.Open(filename)
begin
  sheet = book.WorkSheets(1)
  sheet.UsedRange.Rows.each do |row|
    record = []
    row.Columns.each do |cell|
      record << cell.Text
    end

    unless record[0] == "クラス" then
      puts record.join(",")
      seminarClass = SeminarClass.getInstance(record[0])
      seminarClass.addGroup(record[1], record[2])
    end
  end
  
  SeminarClass.getAllInstance().each do |seminarClass|
    puts "class: " + seminarClass.name.to_s
    seminarClass.groups.each do |group|
      puts "\t" + "group: #{group.name}"
    end
  end
ensure
  book.Close
  xl.Quit
end


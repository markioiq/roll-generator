require 'win32ole'

class ExcelLoader
  attr_reader :symbols
  attr_reader :lines
  def initialize(excelFileName, sheet = 1)
    @symbols = []
    @lines = []

    excel = WIN32OLE.new('Excel.Application')
    filename = getAbsolutePath(excelFileName)
    book = excel.Workbooks.Open(filename: filename, readOnly: true)
    
    begin
      sheet = book.WorkSheets(sheet)
      sheet.UsedRange.Rows.each do |row|
        if (row.Row == 1) then
          row.Columns.each do |column|
            symbols << column.Text.to_sym
          end
        else
          line = Hash.new()
          row.Columns.each do |column|
            line[symbols[column.Column-1]] = column.Text
          end
          lines << line
        end
      end
    ensure
      book.Close
      excel.Quit
    end
  end

  private

  def getAbsolutePath filename
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    return fso.GetAbsolutePathName(filename)
  end
end
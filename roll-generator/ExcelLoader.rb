require 'win32ole'


class ExcelLoader
  attr_reader :symbols
  attr_reader :lines
  def initialize(excelFileName, sheetName = File.basename(excelFileName, ".*"))
    @symbols = []
    @lines = []

    excel = WIN32OLE.new('Excel.Application')
    tmp = excel.DisplayAlerts
    excel.DisplayAlerts = false
    filename = getAbsolutePath(excelFileName)
    book = excel.Workbooks.Open(filename: filename, readOnly: true)
    
    begin
      sheet = book.WorkSheets(sheetName)
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
      excel.DisplayAlerts = tmp
      excel.Quit
    end
  end

  private

  def getAbsolutePath filename
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    return fso.GetAbsolutePathName(filename)
  end
end
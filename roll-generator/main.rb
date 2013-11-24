# vim:set fileencoding=Windows-31J:

MAIN_DIR = File.dirname(File.expand_path(__FILE__))
$LOAD_PATH << MAIN_DIR

require 'win32ole'

#require 'SeatManager'
require 'ExcelLoader'

module Excel
end

excel = WIN32OLE.new('Excel.Application')
excel.Quit
WIN32OLE.const_load(excel, Excel)

def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  return fso.GetAbsolutePathName(filename)
end

def printHash hash
  str = "["
  hash.each_pair {|key, value|
    str += key.to_s + "=>"
    if (value.instance_of?(Hash)) then
      str += printHash(value).to_s
    else
      str += value.to_s
    end
    str += ", "
  }
  str += "]"

  #hash.inject("[") {|result, item|
  #  p item
  #}

  return str
end

seats = []
seatLoader = ExcelLoader.new("./座席番号・定員.xls")
seatLoader.lines.each do |line|
  numOfSeats = line[:定員].to_i
  for i in 1..numOfSeats
    seat = Hash.new
    seat[:クラス] = line[:クラス]
    seat[:グループ] = line[:グループ]
    seat[:番号] = i
    seats << seat
  end
end

#classNames = []
#seatLoader.lines.each do |line|
#  unless (classNames.include?(line[:クラス])) then
#    classNames << line[:クラス]
#  end
#end

traineeLoader = ExcelLoader.new("./受講者リスト.xls")
trainees = traineeLoader.lines()

pastRolls = Dir.glob("[^~]*出欠表.xls")
weight = 1
pastCosts = Hash.new
pastRolls.each do |pastRoll|
  if (pastRoll == "出欠表.xls") then
    next
  end

  pastTraineeLoader = ExcelLoader.new("./受講者リスト.xls")
  pastTrainees = pastTraineeLoader.lines()

  (0..pastTrainees.size-1).to_a.combination(2) {|i, j|
    a = pastTrainees[i]
    b = pastTrainees[j]

    key = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    unless pastCosts.has_key?(key)
      pastCosts[key] = 0
    end

    if (a[:クラス] == b[:クラス] and a[:グループ] == b[:グループ]) then
      pastCosts[key] += 1 * weight
    end
  }

  weight *= 2
end

bestAssignment = []
bestTotalCost = -1
100.times { |n|
  newSeats = seats.shuffle()
  assignedTrainees = []
  trainees.each_with_index {|trainee, index|
    assignedTrainees << trainee.merge({"座席".to_sym => newSeats[index]})
  }

  costs = Hash.new
  (0..assignedTrainees.size-1).to_a.combination(2) {|i, j|
    a = assignedTrainees[i]
    b = assignedTrainees[j]

    cost = 0
    unless (a[:座席][:クラス] == b[:座席][:クラス] and a[:座席][:グループ] == b[:座席][:グループ])
      costs[[i,j]] = cost
      next
    end

    cost += 2048 if (a[:性別] == b[:性別])
    cost += 16 if (a[:係] == b[:係])
    cost += 8 if (a[:課] == b[:課])
    cost += 4 if (a[:部] == b[:部])
    cost += 2 if (a[:職種] == b[:職種])
    cost += 2 if (a[:役職] == b[:役職])
    cost += 1 if (a[:年齢層] == b[:年齢層])

    key = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    if (pastCosts.has_key?(key)) then
      cost += pastCosts[key]
    end

    costs[[i,j]] = cost
  }
  totalCost = costs.values().inject {|result, value| result + value }
  #puts "合計コスト(#{n}): " + totalCost.to_s
  if (bestTotalCost < 0 or totalCost < bestTotalCost) then
    bestAssignment = assignedTrainees
    bestTotalCost = totalCost
  end
}

xl = WIN32OLE.new('Excel.Application')
rollBook = xl.Workbooks.add(getAbsolutePath("./出欠表.xlt"))
seatBook = xl.Workbooks.open(getAbsolutePath("./座席番号・定員.xls"))
begin
  #xl.visible = true
  #rollBook.activate
  #puts xl.Application.Dialogs(Excel::XlDialogSaveAs).show()
  tmp = xl.DisplayAlerts
  xl.DisplayAlerts = false
  rollBook.saveAs(getAbsolutePath('./出欠表.xls'), Excel::XlWorkbookNormal)
  #puts xl.Application.GetSaveAsFilename(getAbsolutePath("."), "座席表Excelファイル, *.xlsx;*.xls")
  xl.DisplayAlerts = tmp

  sheet = rollBook.WorkSheets('出欠表')
  bestAssignment.each_with_index { |trainee, index|
    sheet.Rows(index+2).Columns(1).value = index
    sheet.Rows(index+2).Columns(2).value = trainee[:ID]
    sheet.Rows(index+2).Columns(3).value = trainee[:部]
    sheet.Rows(index+2).Columns(4).value = trainee[:課]
    sheet.Rows(index+2).Columns(5).value = trainee[:係]
    sheet.Rows(index+2).Columns(6).value = trainee[:役職]
    sheet.Rows(index+2).Columns(7).value = trainee[:氏] + "　" + trainee[:名]
    sheet.Rows(index+2).Columns(8).value = trainee[:座席][:クラス]
    sheet.Rows(index+2).Columns(9).value = trainee[:座席][:グループ]
    sheet.Rows(index+2).Columns(10).value = trainee[:座席][:番号]

    sheet.Rows(index+2).Columns(13).value = trainee[:性別]
  }
  rollBook.save

  puts 'aaa'
  #newSheat = rollBook.WorkSheets.add
  #newSheat.name = 'クラス11'
  addr = seatBook.WorkSheets(2).copy("after"=>sheet)
  xl.ActiveSheet.Name = 'クラス11'
  rollBook.save
  puts 'bbb'
ensure
  rollBook.close
  seatBook.close
  xl.Quit
end

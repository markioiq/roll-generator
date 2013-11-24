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
seatLoader = ExcelLoader.new("./���Ȕԍ��E���.xls")
seatLoader.lines.each do |line|
  numOfSeats = line[:���].to_i
  for i in 1..numOfSeats
    seat = Hash.new
    seat[:�N���X] = line[:�N���X]
    seat[:�O���[�v] = line[:�O���[�v]
    seat[:�ԍ�] = i
    seats << seat
  end
end

#classNames = []
#seatLoader.lines.each do |line|
#  unless (classNames.include?(line[:�N���X])) then
#    classNames << line[:�N���X]
#  end
#end

traineeLoader = ExcelLoader.new("./��u�҃��X�g.xls")
trainees = traineeLoader.lines()

pastRolls = Dir.glob("[^~]*�o���\.xls")
weight = 1
pastCosts = Hash.new
pastRolls.each do |pastRoll|
  if (pastRoll == "�o���\.xls") then
    next
  end

  pastTraineeLoader = ExcelLoader.new("./��u�҃��X�g.xls")
  pastTrainees = pastTraineeLoader.lines()

  (0..pastTrainees.size-1).to_a.combination(2) {|i, j|
    a = pastTrainees[i]
    b = pastTrainees[j]

    key = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    unless pastCosts.has_key?(key)
      pastCosts[key] = 0
    end

    if (a[:�N���X] == b[:�N���X] and a[:�O���[�v] == b[:�O���[�v]) then
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
    assignedTrainees << trainee.merge({"����".to_sym => newSeats[index]})
  }

  costs = Hash.new
  (0..assignedTrainees.size-1).to_a.combination(2) {|i, j|
    a = assignedTrainees[i]
    b = assignedTrainees[j]

    cost = 0
    unless (a[:����][:�N���X] == b[:����][:�N���X] and a[:����][:�O���[�v] == b[:����][:�O���[�v])
      costs[[i,j]] = cost
      next
    end

    cost += 2048 if (a[:����] == b[:����])
    cost += 16 if (a[:�W] == b[:�W])
    cost += 8 if (a[:��] == b[:��])
    cost += 4 if (a[:��] == b[:��])
    cost += 2 if (a[:�E��] == b[:�E��])
    cost += 2 if (a[:��E] == b[:��E])
    cost += 1 if (a[:�N��w] == b[:�N��w])

    key = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    if (pastCosts.has_key?(key)) then
      cost += pastCosts[key]
    end

    costs[[i,j]] = cost
  }
  totalCost = costs.values().inject {|result, value| result + value }
  #puts "���v�R�X�g(#{n}): " + totalCost.to_s
  if (bestTotalCost < 0 or totalCost < bestTotalCost) then
    bestAssignment = assignedTrainees
    bestTotalCost = totalCost
  end
}

xl = WIN32OLE.new('Excel.Application')
rollBook = xl.Workbooks.add(getAbsolutePath("./�o���\.xlt"))
seatBook = xl.Workbooks.open(getAbsolutePath("./���Ȕԍ��E���.xls"))
begin
  #xl.visible = true
  #rollBook.activate
  #puts xl.Application.Dialogs(Excel::XlDialogSaveAs).show()
  tmp = xl.DisplayAlerts
  xl.DisplayAlerts = false
  rollBook.saveAs(getAbsolutePath('./�o���\.xls'), Excel::XlWorkbookNormal)
  #puts xl.Application.GetSaveAsFilename(getAbsolutePath("."), "���ȕ\Excel�t�@�C��, *.xlsx;*.xls")
  xl.DisplayAlerts = tmp

  sheet = rollBook.WorkSheets('�o���\')
  bestAssignment.each_with_index { |trainee, index|
    sheet.Rows(index+2).Columns(1).value = index
    sheet.Rows(index+2).Columns(2).value = trainee[:ID]
    sheet.Rows(index+2).Columns(3).value = trainee[:��]
    sheet.Rows(index+2).Columns(4).value = trainee[:��]
    sheet.Rows(index+2).Columns(5).value = trainee[:�W]
    sheet.Rows(index+2).Columns(6).value = trainee[:��E]
    sheet.Rows(index+2).Columns(7).value = trainee[:��] + "�@" + trainee[:��]
    sheet.Rows(index+2).Columns(8).value = trainee[:����][:�N���X]
    sheet.Rows(index+2).Columns(9).value = trainee[:����][:�O���[�v]
    sheet.Rows(index+2).Columns(10).value = trainee[:����][:�ԍ�]

    sheet.Rows(index+2).Columns(13).value = trainee[:����]
  }
  rollBook.save

  puts 'aaa'
  #newSheat = rollBook.WorkSheets.add
  #newSheat.name = '�N���X11'
  addr = seatBook.WorkSheets(2).copy("after"=>sheet)
  xl.ActiveSheet.Name = '�N���X11'
  rollBook.save
  puts 'bbb'
ensure
  rollBook.close
  seatBook.close
  xl.Quit
end

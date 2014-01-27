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
  return str
end

if (ARGV.size() > 0)
  dir = File.dirname(ARGV[0]) + "/"
else
  dir = "./"
end

# �ݒ�t�@�C����ǂݍ���
costTable = Hash.new
configLoader = ExcelLoader.new(dir + "/�ݒ�.xls");
configLoader.lines().each do |line|
  key = line[:����].to_sym
  costTable[key] = []
  configLoader.symbols().each do |sym|
    if (sym == :����) then
      next
    end
    costTable[key] << line[sym].to_i
  end
end
#puts printHash(costTable)

# �e�N���X�E�e�O���[�v�̒��������Ȕԍ��������ȃI�u�W�F�N�g���쐬����B
seats = []
seatLoader = ExcelLoader.new(dir + "/���Ȕԍ��E���.xls")
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

# ��u�҃t�@�C������A��u�҃I�u�W�F�N�g���쐬����
traineeLoader = ExcelLoader.new(dir + "/��u�҃��X�g.xls")
trainees = traineeLoader.lines()

# �ߋ��̍��ȕ\����A�����O���[�v�ɂȂ������Ƃ������u���ԂɁA�R�X�g����悹����B
pastRolls = Dir.glob("[^~]*�o���\*.xls")
weight = 1
pastCosts = Hash.new
pastRolls.reverse.each_with_index do |pastRoll, index|
  if (pastRoll == "�o���\.xls") then
    next
  end

  pastTraineeLoader = ExcelLoader.new(dir + "/��u�҃��X�g.xls")
  pastTrainees = pastTraineeLoader.lines()

  (0..pastTrainees.size-1).to_a.combination(2) {|i, j|
    a = pastTrainees[i]
    b = pastTrainees[j]

    key = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    unless pastCosts.has_key?(key)
      pastCosts[key] = 0
    end

    if (a[:�N���X] == b[:�N���X] and a[:�O���[�v] == b[:�O���[�v]) then
      pastCosts[key] += costTable[:�ߋ�][index]
    end
  }

end

# 100�p�^�[���̍��ȕ\�𐶐����A�����Ƃ����R�X�g�̏��Ȃ����̂��̗p����
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

    costTable.keys().each do |key|
      next if (key == :�ߋ�)

      cost += costTable[key.to_sym][0] if (a[key.to_sym] == b[key.to_sym])
    end

    pairKey = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    if (pastCosts.has_key?(pairKey)) then
      cost += pastCosts[pairKey]
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

# �̗p�������Ȕz�u���t�@�C���ɏ����o��
targetFileLoader = ExcelLoader.new(dir + "/�o���\.xlt")
targetFileSymbols = targetFileLoader.symbols

xl = WIN32OLE.new('Excel.Application')
rollBook = xl.Workbooks.add(getAbsolutePath(dir + "/�o���\.xlt"))
#seatBook = xl.Workbooks.open(getAbsolutePath("./���Ȕԍ��E���.xls"))
begin
  #xl.visible = true
  #rollBook.activate
  #puts xl.Application.Dialogs(Excel::XlDialogSaveAs).show()
  tmp = xl.DisplayAlerts
  xl.DisplayAlerts = false
  rollBook.saveAs(getAbsolutePath(dir + '/�o���\.xls'), Excel::XlWorkbookNormal)
  #puts xl.Application.GetSaveAsFilename(getAbsolutePath("."), "���ȕ\Excel�t�@�C��, *.xlsx;*.xls")
  xl.DisplayAlerts = tmp

  sheet = rollBook.WorkSheets('�o���\')
  bestAssignment.each_with_index { |trainee, lineNumber|
    targetFileSymbols.each_with_index { |symbol, columnNumber|
      if (symbol == "No.".to_sym)
        sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = lineNumber + 1
        next
      end
      
      # ���Ȕԍ����o��
      if ([:�N���X, :�O���[�v, :�ԍ�].include?(symbol))
        sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = trainee[:����][symbol]
        next
      end
      
      if (:���Ȕԍ� == symbol)
        sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = 
          trainee[:����][:�N���X].to_s + "-" + trainee[:����][:�O���[�v].to_s + "-" + trainee[:����][:�ԍ�].to_s 
        next
      end

      # �������o��      
      sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = trainee[symbol]
    }
  }
  rollBook.save

  #newSheat = rollBook.WorkSheets.add
  #newSheat.name = '�N���X11'
  #addr = seatBook.WorkSheets(2).copy("after"=>sheet)
  #xl.ActiveSheet.Name = '�N���X11'
  rollBook.save
ensure
  rollBook.close
  # seatBook.close
  xl.Quit
end

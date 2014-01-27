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

# 設定ファイルを読み込む
costTable = Hash.new
configLoader = ExcelLoader.new(dir + "/設定.xls");
configLoader.lines().each do |line|
  key = line[:属性].to_sym
  costTable[key] = []
  configLoader.symbols().each do |sym|
    if (sym == :属性) then
      next
    end
    costTable[key] << line[sym].to_i
  end
end
#puts printHash(costTable)

# 各クラス・各グループの定員から座席番号を持つ座席オブジェクトを作成する。
seats = []
seatLoader = ExcelLoader.new(dir + "/座席番号・定員.xls")
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

# 受講者ファイルから、受講者オブジェクトを作成する
traineeLoader = ExcelLoader.new(dir + "/受講者リスト.xls")
trainees = traineeLoader.lines()

# 過去の座席表から、同じグループになったことがある受講生間に、コストを上乗せする。
pastRolls = Dir.glob("[^~]*出欠表*.xls")
weight = 1
pastCosts = Hash.new
pastRolls.reverse.each_with_index do |pastRoll, index|
  if (pastRoll == "出欠表.xls") then
    next
  end

  pastTraineeLoader = ExcelLoader.new(dir + "/受講者リスト.xls")
  pastTrainees = pastTraineeLoader.lines()

  (0..pastTrainees.size-1).to_a.combination(2) {|i, j|
    a = pastTrainees[i]
    b = pastTrainees[j]

    key = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    unless pastCosts.has_key?(key)
      pastCosts[key] = 0
    end

    if (a[:クラス] == b[:クラス] and a[:グループ] == b[:グループ]) then
      pastCosts[key] += costTable[:過去][index]
    end
  }

end

# 100パターンの座席表を生成し、もっとも総コストの少ないものを採用する
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

    costTable.keys().each do |key|
      next if (key == :過去)

      cost += costTable[key.to_sym][0] if (a[key.to_sym] == b[key.to_sym])
    end

    pairKey = if a[:ID] <= b[:ID] then [a[:ID], b[:ID]] else [b[:ID], a[:ID]] end
    if (pastCosts.has_key?(pairKey)) then
      cost += pastCosts[pairKey]
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

# 採用した座席配置をファイルに書き出す
targetFileLoader = ExcelLoader.new(dir + "/出欠表.xlt")
targetFileSymbols = targetFileLoader.symbols

xl = WIN32OLE.new('Excel.Application')
rollBook = xl.Workbooks.add(getAbsolutePath(dir + "/出欠表.xlt"))
#seatBook = xl.Workbooks.open(getAbsolutePath("./座席番号・定員.xls"))
begin
  #xl.visible = true
  #rollBook.activate
  #puts xl.Application.Dialogs(Excel::XlDialogSaveAs).show()
  tmp = xl.DisplayAlerts
  xl.DisplayAlerts = false
  rollBook.saveAs(getAbsolutePath(dir + '/出欠表.xls'), Excel::XlWorkbookNormal)
  #puts xl.Application.GetSaveAsFilename(getAbsolutePath("."), "座席表Excelファイル, *.xlsx;*.xls")
  xl.DisplayAlerts = tmp

  sheet = rollBook.WorkSheets('出欠表')
  bestAssignment.each_with_index { |trainee, lineNumber|
    targetFileSymbols.each_with_index { |symbol, columnNumber|
      if (symbol == "No.".to_sym)
        sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = lineNumber + 1
        next
      end
      
      # 座席番号を出力
      if ([:クラス, :グループ, :番号].include?(symbol))
        sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = trainee[:座席][symbol]
        next
      end
      
      if (:座席番号 == symbol)
        sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = 
          trainee[:座席][:クラス].to_s + "-" + trainee[:座席][:グループ].to_s + "-" + trainee[:座席][:番号].to_s 
        next
      end

      # 属性を出力      
      sheet.Rows(lineNumber+2).Columns(columnNumber+1).value = trainee[symbol]
    }
  }
  rollBook.save

  #newSheat = rollBook.WorkSheets.add
  #newSheat.name = 'クラス11'
  #addr = seatBook.WorkSheets(2).copy("after"=>sheet)
  #xl.ActiveSheet.Name = 'クラス11'
  rollBook.save
ensure
  rollBook.close
  # seatBook.close
  xl.Quit
end

class Seat
  attr_reader :group
  attr_reader :name
  
  def initialize(group, name)
    @group = group
    @name = name
  end
  
  def path
    string = ""
    string += @group.seminarClass.name
    string += "/" + @group.name
    string += "/" + name
  end
end
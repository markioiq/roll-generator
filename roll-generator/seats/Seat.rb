class Seat
  attr_reader :group
  attr_reader :name
  
  def initialize(group, name)
    @group = group
    @name = name
  end
end
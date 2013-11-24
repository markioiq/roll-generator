require 'SeminarGroup'

class SeminarClass
  attr_reader :name
  attr_reader :groups
  def initialize(name)
    @name = name
    @groups = Hash.new { |hash, key|
      newInstance = SeminarGroup.new(self, key)
      hash[key] = newInstance
    }

  end

  def getAllGroups
    @groups.values()
  end

  def each_seat(&block)
    @groups.values().each do |group|
      group.each_seat(&block)
    end

  end

end
require 'SeminarGroup'

class SeminarClass
  attr_reader :name
  attr_reader :groups
  
  @@instances = Hash.new([])
  
  class << self
    def getInstance(name)
      key = name.to_s.to_sym
      if @@instances.has_key?(key) then
        return @@instances[key]
      else
        newInstance = self.new(name)
        @@instances[key] = newInstance
        return newInstance
      end
    end
    
    def getAllInstance
      return @@instances.values()
    end
  end
  
  def initialize(name)
    @name = name
    @groups = []
  end
  
  def addGroup name
    group = SeminarGroup.new(name)
    @groups << group
  end
    
end
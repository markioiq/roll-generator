require 'SeminarClass'
require 'singleton'

class SeatManager

  include Singleton
  def initialize
    @seminarClasses = Hash.new { |hash, key|
      newInstance = SeminarClass.new(key)
      hash[key] = newInstance
    }
  end

  def getSeminarClass(name)
    key = name.to_s
    return @seminarClasses[key]
  end

  def getSeminarClasses
    return @seminarClasses.values()
  end

  def makeSeat(seminarClassName, seminarGroupName, numberOfSeat)
    seminarClass = @seminarClasses[seminarClassName]
    seminarGroup = seminarClass.groups[seminarGroupName]
    seminarGroup.addSeats(numberOfSeat)
  end

  def each_seat(&block)
    @seminarClasses.values().each do |seminarClass|
      seminarClass.each_seat(&block)
    end
  end

end
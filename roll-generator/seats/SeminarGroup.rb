require 'Seat'

class SeminarGroup
  attr_reader :seminarClass
  attr_reader :name
  attr_reader :seats
  def initialize(seminarClass, name)
    @seminarClass = seminarClass
    @name = name
    @seats = []
  end

  def addSeats(numberOfSeats)
    numberOfSeats.times do
      seat = Seat.new(self, (seats.size + 1).to_s)
      @seats << seat
    end
  end
  
  def each_seat(&block)
    @seats.each do |seat|
      block.call(seat)
    end
  end
end
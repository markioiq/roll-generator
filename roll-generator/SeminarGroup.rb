require './Seat'

class SeminarGroup
  attr_reader :seminarClass
  attr_reader :name
  attr_reader :seats
  
  def initialize(seminarClass, name, numberOfSeats)
    @seminarClass = seminarClass
    @name = name
    @seats = []
    
    numberOfSeats.to_i.times do
      seat = Seat.new(self, seats.size + 1)
      @seats << seat
    end
  end
end
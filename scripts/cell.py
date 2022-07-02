from enum import Enum

class Color(Enum):
    RED = 1
    WHITE = 2
    
class Cell:
    def __init__(self, id, position_x, position_y, type: Color):
        self.id = id
        self.position_x = position_x
        self.position_y = position_y
        self.type = type
    
    
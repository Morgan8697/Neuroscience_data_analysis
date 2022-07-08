from enum import Enum

class Family(Enum):
    MICROGLIA = 1
    PNN = 2
    
    
class Cell:
    def __init__(self, id, position_x, position_y, family: Family, type, sheet_name):
        self.id = id
        self.position_x = position_x
        self.position_y = position_y
        self.family = family
        self.type = type
        self.sheet_name = sheet_name
    
    
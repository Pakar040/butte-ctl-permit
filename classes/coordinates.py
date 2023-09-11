class Coordinates:
    """
    Represents a position on the planet
    """
    def __init__(self, latitude: float, longitude: float):
        self.latitude = latitude
        self.longitude = longitude

    def distance_from_coordinate(self, other: 'Coordinates'):
        pass

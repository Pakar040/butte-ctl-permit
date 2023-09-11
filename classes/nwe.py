from dataclasses import dataclass, field
from coordinates import Coordinates


class NWEXLSX:
    """
    Represents an NWE XLSX file
    """

    def __init__(self, filepath: str):
        self.pole_data_list = []
        self.filepath = filepath

    def read_xlsx(self):
        pass


@dataclass
class NWEPole:
    """
    Contains the pole data provided by the NWE XLSX file
    """

    pole_number: str
    coordinates: Coordinates
    attachment_list: list = field(default_factory=list)

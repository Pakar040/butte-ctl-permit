from dataclasses import dataclass
from coordinates import Coordinates

class AldenoneXLSX:
    """
    Represents Aldenone Excel file
    """

    def __init__(self, filepath: str):
        self.pole_data_list = []
        self.filepath = filepath

    def read_xlsx(self):
        pass


@dataclass
class AldenonePole:
    """
    Contains the pole data provided by the aldenone XLSX file
    """

    is_ctl_pole: bool
    pole_tag_number: str
    coordinates: Coordinates

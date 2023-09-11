from dataclasses import dataclass, field
from coordinates import Coordinates


class CTLXLSX:
    """
    Represents an CTL XLSX file
    """

    def __init__(self, filepath: str):
        self.pole_data_list = []
        self.filepath = filepath

    def read_xlsx(self):
        pass


@dataclass
class CTLPole:

    pole_number: str
    coordinates: Coordinates
    attachment_list: list = field(default_factory=list)

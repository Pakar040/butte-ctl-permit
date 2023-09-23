from dataclasses import dataclass, field
from openpyxl import load_workbook


class CTLXLSX:
    """
    Represents an CTL XLSX file
    """

    def __init__(self, filepath: str):
        self.pole_data_list = []
        self.filepath = filepath
        self.worksheet = None
        self.column_headers = None
        self._read_xlsx()
        self._get_column_headers()

    def _read_xlsx(self):
        workbook = load_workbook(self.filepath)
        self.worksheet = workbook['Sheet1']

    def _get_column_headers(self):
        column_header_cell_values = self.worksheet[14]
        self.column_headers = [cell.value for cell in column_header_cell_values]


@dataclass
class CTLPole:

    pole_number: str
    attachment_list: list = field(default_factory=list)

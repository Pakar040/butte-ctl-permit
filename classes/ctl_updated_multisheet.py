from dataclasses import dataclass, field
from typing import List
import os
import xlrd
from xlutils.copy import copy  # Replacing openpyxl with xlrd and xlutils import os
import helper.functions as hf


class CTLXLSX:
    """
    Represents an CTL XLSX file
    """

    def __init__(self, filepath: str):
        self.pole_data_list = []
        self.template_filepath = self._get_template_mac_windows()
        self.filepath = filepath
        self.read_workbook = None
        self.write_workbook = None
        self.active_sheet = None
        self._load_template_xls()

    def set_pole_data_list(self, lst: List['CTLPole']):
        self.pole_data_list = lst

    def write_to_cell(self, value: any, row: int, column: int, sheet='Sheet1'):
        sheet = self.write_workbook.get_sheet(sheet)
        sheet.write(row - 1, column - 1, value)

    def save_workbook(self):
        self.write_workbook.save(self.filepath)

    def write_all_data_to_workbook(self):
        start_row = 18
        page = 0
        counter = 0
        for index, ctl_pole in enumerate(self.pole_data_list):
            if counter == 15:
                page += 1

                if page == len(self.read_workbook.sheet_names()):
                    print("OUT OF ROWS!!!")
                    break

                self.active_sheet = self.read_workbook.sheet_names()[page]
                counter = 0

            row = start_row + index - (page * 15)
            self.write_pole_data(row=row, ctl_pole=ctl_pole)
            counter = counter + 1

    def write_pole_data(self, row: int, ctl_pole: 'CTLPole'):
        # ----- temporary ----- #
        self.write_to_cell(value=ctl_pole.pole_number, row=row, column=4, sheet=self.active_sheet)

        # --------------------- #
        self.write_attachment_action(row=row, ctl_pole=ctl_pole)
        self.write_ctl_pole_number(row=row, ctl_pole=ctl_pole)
        self.write_latitude(row=row, ctl_pole=ctl_pole)
        self.write_longitude(row=row, ctl_pole=ctl_pole)
        self.write_pole_type(row=row, ctl_pole=ctl_pole)
        self.write_pole_height_class(row=row, ctl_pole=ctl_pole)
        self.write_attachment_type(row=row, ctl_pole=ctl_pole)
        self.write_new_attachment_height(row=row, ctl_pole=ctl_pole)
        self.write_pole_side(row=row, ctl_pole=ctl_pole)
        self.write_telco_attachment_type(row=row, ctl_pole=ctl_pole)
        self.write_existing_telco_attach_height(row=row, ctl_pole=ctl_pole)
        self.write_lowest_power_secondary(row=row, ctl_pole=ctl_pole)
        self.write_streetlight_mast(row=row, ctl_pole=ctl_pole)
        self.write_streetlight_conductor(row=row, ctl_pole=ctl_pole)
        self.write_lowest_comm_midspan(row=row, ctl_pole=ctl_pole)
        self.write_catv_height(row=row, ctl_pole=ctl_pole)
        self.write_other_height(row=row, ctl_pole=ctl_pole)
        self.write_messenger_size(row=row, ctl_pole=ctl_pole)
        self.write_span_length(row=row, ctl_pole=ctl_pole)
        self.write_make_ready_notes(row=row, ctl_pole=ctl_pole)

    def write_attachment_action(self, row: int, ctl_pole: 'CTLPole'):
        attachment_action = 'O - Overlash'
        attachment_action_column = 2  # B
        self.write_to_cell(value=attachment_action, row=row, column=attachment_action_column, sheet=self.active_sheet)
        self.write_to_cell(value=attachment_action, row=row + 17, column=attachment_action_column,
                           sheet=self.active_sheet)

    def write_ctl_pole_number(self, row: int, ctl_pole: 'CTLPole'):
        ctl_pole_number = ctl_pole.ctl_pole_number
        ctl_pole_column = 6  # F
        self.write_to_cell(value=ctl_pole_number, row=row, column=ctl_pole_column, sheet=self.active_sheet)
        self.write_to_cell(value=ctl_pole_number, row=row + 17, column=ctl_pole_column, sheet=self.active_sheet)

    def write_latitude(self, row: int, ctl_pole: 'CTLPole'):
        latitude = ctl_pole.latitude
        latitude_column = 12  # L
        self.write_to_cell(value=latitude, row=row, column=latitude_column, sheet=self.active_sheet)

    def write_longitude(self, row: int, ctl_pole: 'CTLPole'):
        longitude = ctl_pole.longitude
        longitude_column = 13  # M
        self.write_to_cell(value=longitude, row=row, column=longitude_column, sheet=self.active_sheet)

    def write_pole_type(self, row: int, ctl_pole: 'CTLPole'):
        pole_type = 'Distribution'
        pole_type_column = 14  # N
        self.write_to_cell(value=pole_type, row=row, column=pole_type_column, sheet=self.active_sheet)

    def write_pole_height_class(self, row: int, ctl_pole: 'CTLPole'):
        pole_height_class = ctl_pole.pole_height_class
        pole_height_class_column = 15  # O
        self.write_to_cell(value=pole_height_class, row=row, column=pole_height_class_column, sheet=self.active_sheet)

    def write_attachment_type(self, row: int, ctl_pole: 'CTLPole'):
        communication_make_ready = ctl_pole.communication_make_ready
        is_drop = any(keyword in communication_make_ready.lower() for keyword in ['drop', 'j-hook', 'j hook'])
        if is_drop:
            attachment_type = 'COMD - Comm Drop'
        else:
            attachment_type = 'COFO - Comm Fiber'

        attachment_type_column = 17  # Q
        self.write_to_cell(value=attachment_type, row=row, column=attachment_type_column, sheet=self.active_sheet)

    def write_new_attachment_height(self, row: int, ctl_pole: 'CTLPole'):
        ziply_proposed_height = None
        for attachment_obj in ctl_pole.attachment_list:
            attachment_name = attachment_obj['attachment_name']
            if 'ziply fiber' in attachment_name.lower():
                ziply_proposed_height = attachment_obj['proposed_height']

        new_attachment_height_column = 18  # R
        self.write_to_cell(value=ziply_proposed_height, row=row, column=new_attachment_height_column,
                           sheet=self.active_sheet)

    def write_pole_side(self, row: int, ctl_pole: 'CTLPole'):
        pole_side = 'RD - Road Side'
        pole_side_column = 19  # S
        self.write_to_cell(value=pole_side, row=row, column=pole_side_column, sheet=self.active_sheet)

    def write_telco_attachment_type(self, row: int, ctl_pole: 'CTLPole'):
        telco_attachment_type = 'COML - Comm Main Line'
        telco_attachment_type_column = 21  # U
        self.write_to_cell(value=telco_attachment_type, row=row, column=telco_attachment_type_column,
                           sheet=self.active_sheet)

    def write_existing_telco_attach_height(self, row: int, ctl_pole: 'CTLPole'):
        cell_data = ""
        for attachment_obj in ctl_pole.attachment_list:
            attachment_name = attachment_obj['attachment_name']
            if attachment_name.lower() == 'lumen':
                if self._string_is_empty(cell_data):
                    cell_data = self._get_attachment_height(attachment_obj)
                else:
                    cell_data = cell_data + "  " + self._get_attachment_height(attachment_obj)

        existing_telco_attach_height_column = 22  # V
        self.write_to_cell(value=cell_data, row=row, column=existing_telco_attach_height_column, sheet=self.active_sheet)

    def write_lowest_power_secondary(self, row: int, ctl_pole: 'CTLPole'):
        lowest_power_attachment = None
        lowest_power_height = None
        for attachment_obj in ctl_pole.attachment_list:
            if self._is_power(attachment_obj):
                if lowest_power_attachment is None:
                    lowest_power_attachment = attachment_obj
                elif self._attachment_is_lower(attachment_obj, lowest_power_attachment):
                    lowest_power_attachment = attachment_obj

        if lowest_power_attachment is not None:
            lowest_power_height = self._get_attachment_height(lowest_power_attachment)

        lowest_power_column = 23  # W
        self.write_to_cell(value=lowest_power_height, row=row, column=lowest_power_column,
                           sheet=self.active_sheet)

    def write_streetlight_mast(self, row: int, ctl_pole: 'CTLPole'):
        streetlight_height = None
        for attachment_obj in ctl_pole.attachment_list:
            if self._is_streetlight(attachment_obj):
                streetlight_height = self._get_attachment_height(attachment_obj)

        streetlight_column = 24  # X
        self.write_to_cell(value=streetlight_height, row=row, column=streetlight_column,
                           sheet=self.active_sheet)

    def write_streetlight_conductor(self, row: int, ctl_pole: 'CTLPole'):
        streetlight_drip_height = None
        for attachment_obj in ctl_pole.attachment_list:
            if self._is_streetlight_drip(attachment_obj):
                streetlight_drip_height = self._get_attachment_height(attachment_obj)

        streetlight_drip_column = 25  # Y
        self.write_to_cell(value=streetlight_drip_height, row=row, column=streetlight_drip_column,
                           sheet=self.active_sheet)

    def write_lowest_comm_midspan(self, row: int, ctl_pole: 'CTLPole'):
        lowest_comm_midspan = None
        lowest_comm_height = None
        for midspan_obj in ctl_pole.midspan_list:
            if self._is_comm(midspan_obj):
                if lowest_comm_midspan is None:
                    lowest_comm_midspan = midspan_obj
                elif self._midspan_is_lower(midspan_obj, lowest_comm_midspan):
                    lowest_comm_midspan = midspan_obj

        if lowest_comm_midspan is not None:
            lowest_comm_height = self._get_midspan_height(lowest_comm_midspan)

        lowest_midspan_column = 26  # Z
        self.write_to_cell(value=lowest_comm_height, row=row, column=lowest_midspan_column,
                           sheet=self.active_sheet)

    def write_catv_height(self, row: int, ctl_pole: 'CTLPole'):
        catv_height = None
        for attachment_obj in ctl_pole.attachment_list:
            attachment_name = attachment_obj['attachment_name']
            if 'spectrum catv' in attachment_name.lower() and 'downguy' not in attachment_name.lower():
                catv_height = self._get_attachment_height(attachment_obj)

        catv_height_column = 27  # AA
        self.write_to_cell(value=catv_height, row=row, column=catv_height_column,
                           sheet=self.active_sheet)

    def write_other_height(self, row: int, ctl_pole: 'CTLPole'):  # TDS
        other_height = None
        for attachment_obj in ctl_pole.attachment_list:
            attachment_name = attachment_obj['attachment_name']
            if 'tds' in attachment_name.lower() and 'downguy' not in attachment_name.lower():
                other_height = self._get_attachment_height(attachment_obj)

        other_height_column = 28  # AB
        self.write_to_cell(value=other_height, row=row, column=other_height_column,
                           sheet=self.active_sheet)

    def write_messenger_size(self, row: int, ctl_pole: 'CTLPole'):
        communication_make_ready = ctl_pole.communication_make_ready
        is_drop = any(keyword in communication_make_ready.lower() for keyword in ['drop', 'j-hook', 'j hook'])
        if is_drop:
            messenger_size = None
        else:
            messenger_size = '10M'

        messenger_size_column = 9  # I
        row = row + 17  # bottom section
        self.write_to_cell(value=messenger_size, row=row, column=messenger_size_column, sheet=self.active_sheet)

    def write_span_length(self, row: int, ctl_pole: 'CTLPole'):
        span_length = ctl_pole.span_length
        span_length_column = 16  # P
        row = row + 17  # bottom section
        self.write_to_cell(value=span_length, row=row, column=span_length_column, sheet=self.active_sheet)

    def write_make_ready_notes(self, row: int, ctl_pole: 'CTLPole'):
        communication_make_ready = ctl_pole.communication_make_ready
        utility_make_ready = ctl_pole.utility_make_ready
        make_ready_notes = f"{communication_make_ready}\n{utility_make_ready}"
        make_ready_notes_column = 20  # T
        row = row + 17  # bottom section
        self.write_to_cell(value=make_ready_notes, row=row, column=make_ready_notes_column, sheet=self.active_sheet)

    @staticmethod
    def _is_comm(attachment_obj):
        name = attachment_obj['attachment_name']
        if 'ziply' in name.lower():
            return True
        elif 'spectrum' in name.lower():
            return True
        elif 'lumen' in name.lower():
            return True
        else:
            return False

    @staticmethod
    def _is_streetlight_drip(attachment_obj):
        name = attachment_obj['attachment_name']
        if 'street light drip' in name.lower():
            return True
        elif 'streetlight drip' in name.lower():
            return True
        else:
            return False

    @staticmethod
    def _is_streetlight(attachment_obj):
        name = attachment_obj['attachment_name']
        if 'street light' in name.lower() and 'drip' not in name.lower():
            return True
        elif 'streetlight' in name.lower() and 'drip' not in name.lower():
            return True
        else:
            return False

    def _attachment_is_lower(self, attachment_obj, other):
        attachment_height = self._get_attachment_height(attachment_obj)
        other_height = self._get_attachment_height(other)
        attachment_height_inches = int(hf.convert_to_inches(attachment_height))
        other_height_inches = int(hf.convert_to_inches(other_height))
        if attachment_height_inches < other_height_inches:
            return True
        else:
            return False

    @staticmethod
    def _get_attachment_height(attachment_obj):
        proposed_height = attachment_obj['proposed_height']
        existing_height = attachment_obj['existing_height']
        if proposed_height is not None:
            return proposed_height
        else:
            return existing_height

    def _midspan_is_lower(self, midspan_obj, other):
        midspan_height = self._get_midspan_height(midspan_obj)
        other_height = self._get_midspan_height(other)
        midspan_height_inches = int(hf.convert_to_inches(midspan_height))
        other_height_inches = int(hf.convert_to_inches(other_height))
        if midspan_height_inches < other_height_inches:
            return True
        else:
            return False

    @staticmethod
    def _get_midspan_height(midspan_obj):
        midspan_height = midspan_obj['midspan_height']
        return midspan_height

    @staticmethod
    def _is_power(attachment_obj):
        name = attachment_obj['attachment_name']
        if 'nwe' in name.lower():
            return True
        elif 'secondary' in name.lower():
            return True
        elif 'power' in name.lower():
            return True
        elif 'crossarm brace' in name.lower():
            return True
        elif 'drip loop' in name.lower():
            return True
        else:
            return False

    @staticmethod
    def _string_is_empty(string: str):
        if string == "":
            return True
        else:
            return False

    @staticmethod
    def _is_lumen_attachment(attachment_name: str):
        if attachment_name.lower() == 'lumen':
            return True
        else:
            return False

    def _load_template_xls(self):
        self.read_workbook = xlrd.open_workbook(filename=self.template_filepath)
        self.write_workbook = copy(self.read_workbook)
        self.active_sheet = self.read_workbook.sheet_names()[0]

    @staticmethod
    def _get_template_mac_windows():
        # Get the directory of the current script
        current_directory = os.path.dirname(os.path.abspath(__file__))

        # Go up one directory to the project folder
        project_directory = os.path.dirname(current_directory)

        # Construct the full path to the Excel file for Windows and Mac
        excel_path = os.path.join(project_directory, 'templates', 'ctl_permit_template.xls')

        return excel_path


@dataclass
class CTLPole:

    pole_number: str
    ctl_pole_number: str
    latitude: str
    longitude: str
    pole_height_class: str
    span_length: str
    communication_make_ready: str
    utility_make_ready: str
    attachment_list: list = field(default_factory=list)
    midspan_list: list = field(default_factory=list)

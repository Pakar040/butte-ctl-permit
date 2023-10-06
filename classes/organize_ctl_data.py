from typing import List, Dict
from classes.ctl import CTLPole
from classes.nwe import NWEXLSX
from classes.aldenone import AldenoneXLSX
import helper.functions as hf


class CTLPoleListFactory:

    def __init__(self, nwe_xlsx_obj: NWEXLSX, aldenone_xlsx_obj: AldenoneXLSX):
        self.nwe_xlsx_obj = nwe_xlsx_obj
        self.aldenone_xlsx_obj = aldenone_xlsx_obj
        self.ctl_data = self._combine_nwe_ctl_data()
        self.ctl_pole_list = [CTLPole]

    def __repr__(self):
        string = ""
        for pole_data in self.ctl_data:
            string += f"{pole_data}" + "\n"
        return string

    def _combine_nwe_ctl_data(self):
        aldenone_data = self.aldenone_xlsx_obj.data
        ctl_data = []
        for aldenone_pole in aldenone_data:
            nwe_pole = self._find_aldenones_pair_from_nwe_data(aldenone_pole)
            ctl_pole = {**aldenone_pole, **nwe_pole}
            ctl_data.append(ctl_pole)

        return ctl_data

    def _find_aldenones_pair_from_nwe_data(self, aldenone_pole: Dict):
        aldenone_pole_number = aldenone_pole['Pole Number']
        for nwe_pole in self.nwe_xlsx_obj.data:
            nwe_pole_number = nwe_pole['Pole Number']
            is_same_pole_number = aldenone_pole_number == nwe_pole_number
            if is_same_pole_number:
                return nwe_pole

    def create_ctl_poles_list(self):
        ctl_pole_list = []
        for pole in self.ctl_data:
            pole_obj = self._create_single_ctl_pole(pole_data=pole)
            ctl_pole_list.append(pole_obj)

        self.ctl_pole_list = ctl_pole_list
        return ctl_pole_list

    @staticmethod
    def _create_single_ctl_pole(pole_data):
        arguments = {
            'pole_number': pole_data['Pole Number'],
            'ctl_pole_number': pole_data['CTL Tags'],
            'latitude': pole_data['Latitude'],
            'longitude': pole_data['Longitude'],
            'pole_height_class': pole_data['Pole Ht/ Class'],
            'span_length': pole_data['Span Length'],
            'communication_make_ready': pole_data['Communication Make Ready'],
            'utility_make_ready': pole_data['Utility Make Ready'],
            'attachment_list': CTLPoleListFactory._create_attachment_list(pole_data),
            'midspan_list': CTLPoleListFactory._create_midspan_list(pole_data),
        }

        return CTLPole(**arguments)

    @staticmethod
    def _create_attachment_list(pole_data: Dict):
        attachment_list = []
        for index, attachment in enumerate(pole_data['Attacher Description']):
            attachment_obj = {
                'attachment_name': hf.remove_after_open_parenthesis(attachment),
                'existing_height': pole_data['Attachment Height Existing'][index],
                'proposed_height': pole_data['Attachment Height Proposed'][index]
            }

            is_duplicate = hf.check_duplicate_dict(attachment_obj, attachment_list)
            has_height_value = attachment_obj['existing_height'] is not None or attachment_obj['proposed_height'] is not None
            if not is_duplicate and has_height_value:
                attachment_list.append(attachment_obj)

        return attachment_list

    @staticmethod
    def _create_midspan_list(pole_data: Dict):
        midspan_list = []
        for index, midspan in enumerate(pole_data['Attacher Description']):
            midspan_obj = {
                'attachment_name': hf.remove_after_open_parenthesis(midspan),
                'midspan_height': pole_data['Midspan Existing'][index],
            }

            is_duplicate = hf.check_duplicate_dict(midspan_obj, midspan_list)
            has_height_value = midspan_obj['midspan_height'] is not None
            if not is_duplicate and has_height_value:
                midspan_list.append(midspan_obj)

        return midspan_list


class CTLPoleFactory:

    def __init__(self):
        self.ctl_pole: CTLPole

    def create_ctl_pole_using_aldenone_nwe(self):
        pass

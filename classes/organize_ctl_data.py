from aldenone import AldenonePole
from nwe import NWEPole
from ctl import CTLPole


class CTLPoleListFactory:

    def __init__(self):
        self.aldenone_pole_list = [AldenonePole]
        self.nwe_pole_list = [NWEPole]
        self.ctl_pole_list = [CTLPole]

    def create_ctl_poles(self):
        pass


class CTLPoleFactory:

    def __init__(self):
        self.aldenone_pole: AldenonePole
        self.nwe_pole: NWEPole
        self.ctl_pole: CTLPole

    def create_ctl_pole_using_aldenone_nwe(self):
        pass

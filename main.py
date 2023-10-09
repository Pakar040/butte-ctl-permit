from classes.nwe import NWEXLSX
from classes.aldenone import AldenoneXLSX
from classes.organize_ctl_data import CTLPoleListFactory
from classes.ctl import CTLXLSX
from classes.app import App
from tkinter import Tk


def main():
    root = Tk()
    app = App(root)
    root.mainloop()
    # nwe_xlsx = NWEXLSX('/Users/alekkariniemi/Library/CloudStorage/GoogleDrive-alek@bookereng.com/Shared drives/2023 '
    #                    'Booker Engineering/CLIENTS/TRACK UTILITIES/Butte/O-Calc/Resources/Alek CTL Automation/Permit '
    #                    '1.xlsx')
    # aldenone_xlsx = AldenoneXLSX('/Users/alekkariniemi/Library/CloudStorage/GoogleDrive-alek@bookereng.com/Shared '
    #                              'drives/2023 Booker Engineering/CLIENTS/TRACK UTILITIES/Butte/O-Calc/Resources/Alek '
    #                              'CTL Automation/1014 Pole info/CTL TAGS.xlsx')
    # ctl_pole_data_factory_obj = CTLPoleListFactory(nwe_xlsx, aldenone_xlsx)
    # ctl_pole_list = ctl_pole_data_factory_obj.create_ctl_poles_list()
    # ctl_xlsx = CTLXLSX(filepath='output/output.xlsx')
    # ctl_xlsx.set_pole_data_list(ctl_pole_list)
    # ctl_xlsx.write_all_data_to_workbook()
    # ctl_xlsx.save_workbook()


if __name__ == "__main__":
    main()

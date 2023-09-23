from classes.nwe import NWEXLSX
from classes.aldenone import AldenoneXLSX


def main():
    nwe_xlsx = NWEXLSX('/Users/alekkariniemi/Library/CloudStorage/GoogleDrive-alek@bookereng.com/Shared drives/2023 '
                       'Booker Engineering/CLIENTS/TRACK UTILITIES/Butte/O-Calc/Resources/Alek CTL Automation/Permit '
                       '1.xlsx')

    print(nwe_xlsx)

    aldenone_xlsx = AldenoneXLSX('/Users/alekkariniemi/Library/CloudStorage/GoogleDrive-alek@bookereng.com/Shared '
                                 'drives/2023 Booker Engineering/CLIENTS/TRACK UTILITIES/Butte/O-Calc/Resources/Alek '
                                 'CTL Automation/1014 Pole info/CTL TAGS.xlsx')

    print(aldenone_xlsx)


if __name__ == "__main__":
    main()

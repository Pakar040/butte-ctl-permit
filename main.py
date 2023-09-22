from classes.nwe import NWEXLSX


def main():
    nwe_xlsx = NWEXLSX('/Users/alekkariniemi/Library/CloudStorage/GoogleDrive-alek@bookereng.com/Shared drives/2023 '
                       'Booker Engineering/CLIENTS/TRACK UTILITIES/Butte/O-Calc/Resources/Alek CTL Automation/Permit '
                       '1.xlsx')

    nwe_permit_data = nwe_xlsx.extract_data()
    for index, pole in enumerate(nwe_permit_data):
        print(pole)


if __name__ == "__main__":
    main()

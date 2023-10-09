from tkinter import Tk, Label, Button, filedialog
import os
from classes.nwe import NWEXLSX
from classes.aldenone import AldenoneXLSX
from classes.organize_ctl_data import CTLPoleListFactory
from classes.ctl import CTLXLSX

class App:
    def __init__(self, root):
        self.root = root
        root.title("File Selector")
        root.geometry("400x300")  # Making the Tkinter box bigger

        self.nwe_file_path = ""
        self.aldenone_file_path = ""
        self.output_folder_path = ""

        self.label = Label(root, text="Select files and output folder")
        self.label.pack()

        self.nwe_label = Label(root, text="NWEXLSX File: Not selected")
        self.nwe_label.pack()

        self.select_nwe_button = Button(root, text="Select NWEXLSX", command=self.select_nwe_file)
        self.select_nwe_button.pack()

        self.aldenone_label = Label(root, text="AldenoneXLSX File: Not selected")
        self.aldenone_label.pack()

        self.select_aldenone_button = Button(root, text="Select AldenoneXLSX", command=self.select_aldenone_file)
        self.select_aldenone_button.pack()

        self.output_folder_label = Label(root, text="Output Folder: Not selected")
        self.output_folder_label.pack()

        self.select_output_folder_button = Button(root, text="Select Output Folder", command=self.select_output_folder)
        self.select_output_folder_button.pack()

        self.run_button = Button(root, text="Run", command=self.execute_logic)
        self.run_button.pack()

    def select_nwe_file(self):
        self.nwe_file_path = filedialog.askopenfilename(title="Select NWEXLSX file")
        self.nwe_label.config(text=f"NWEXLSX File: {os.path.basename(self.nwe_file_path)}")

    def select_aldenone_file(self):
        self.aldenone_file_path = filedialog.askopenfilename(title="Select AldenoneXLSX file")
        self.aldenone_label.config(text=f"AldenoneXLSX File: {os.path.basename(self.aldenone_file_path)}")

    def select_output_folder(self):
        self.output_folder_path = filedialog.askdirectory(title="Select Output Folder")
        self.output_folder_label.config(text=f"Output Folder: {os.path.basename(self.output_folder_path)}")

    def execute_logic(self):
        nwe_xlsx = NWEXLSX(self.nwe_file_path)
        aldenone_xlsx = AldenoneXLSX(self.aldenone_file_path)
        ctl_pole_data_factory_obj = CTLPoleListFactory(nwe_xlsx, aldenone_xlsx)
        ctl_pole_list = ctl_pole_data_factory_obj.create_ctl_poles_list()
        ctl_xlsx = CTLXLSX(filepath=os.path.join(self.output_folder_path, 'output.xlsx'))
        ctl_xlsx.set_pole_data_list(ctl_pole_list)
        ctl_xlsx.write_all_data_to_workbook()
        ctl_xlsx.save_workbook()

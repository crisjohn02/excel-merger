import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askdirectory
import pandas as pd
import os
import urllib.parse
import glob


def select_file():
    filename = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])
    return filename
    # return askdirectory()


def extract_url(url, keys):
    res = urllib.parse.parse_qs(url)
    return {key: value[0] for (key, value) in res.items() if key in keys}


def run_merge():
    base_file = select_file()
    base_df1 = pd.read_excel(base_file, sheet_name="Computed Scores")
    base_df2 = pd.read_excel(base_file, sheet_name="IQ Scores")
    base_df3 = pd.read_excel(base_file, sheet_name="Counts")

    source_path = askdirectory()
    files = glob.glob(source_path + "/*.xlsx")
    for file in files:
        base1 = base_df1.copy()
        base2 = base_df2.copy()
        base3 = base_df3.copy()
        target_dir = source_path + "/results"
        if not os.path.exists(target_dir):
            os.makedirs(target_dir)
        save_path = target_dir + "/" + os.path.basename(file)
        writer = pd.ExcelWriter(save_path)
        base1 = pd.concat([base1, pd.read_excel(file, sheet_name="Computed Scores")])
        base2 = pd.concat([base2, pd.read_excel(file, sheet_name="IQ Scores")])
        base3 = pd.concat([base3, pd.read_excel(file, sheet_name="Counts")])

        base1.to_excel(writer, sheet_name="Computed Scores", index=False)
        base2.to_excel(writer, sheet_name="IQ Scores", index=False)
        base3.to_excel(writer, sheet_name="Counts", index=False)
        writer.save()


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.geometry("760x400")
        self.title("Excel Appender")
        # self.resizable(0, 0)

        # configure the grid
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)

        # variables
        self.source_filename = None
        self.target_filename = None
        self.source_dataframe = None
        self.target_dataframe = None

        self.source_base_column = None
        self.target_base_column = None
        self.target_position = None
        self.url_keys = None

        # widgets
        self.source_entry = None
        self.source_excluded_columns = None
        self.target_entry = None
        self.target_excluded_columns = None

        # Initialize and create the UX widgets
        self.create_widgets()

    def create_widgets(self):

        ttk.Label(self, text="Step 1: Click the button below and Select the base Excel file").grid(column=0, row=0,
                                                                                                   sticky=tk.NW, padx=5,
                                                                                                   pady=5)
        ttk.Label(self, text="Step 2: Select the folder containing the unaligned excel files").grid(column=0, row=1,
                                                                                                    sticky=tk.NW,
                                                                                                    padx=5, pady=5)
        ttk.Label(self, text="Step 3: The processed files will be in results folders within your selected folder").grid(
            column=0, row=2, sticky=tk.NW, padx=5, pady=5)

        login_button = ttk.Button(self, text="Align AST Tabular Excel", command=run_merge)
        login_button.grid(column=0, row=3, sticky=tk.NW, padx=5, pady=5)

    def select_source(self):
        self.source_filename = select_file()
        self.source_entry.delete(0, END)
        self.source_entry.insert(0, os.path.basename(self.source_filename))

    def select_target(self):
        self.target_filename = select_file()
        self.target_entry.delete(0, END)
        self.target_entry.insert(0, os.path.basename(self.target_filename))


if __name__ == "__main__":
    app = App()
    app.mainloop()

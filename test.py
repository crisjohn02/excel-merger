# from tkinter.filedialog import askopenfilenames
# import os
# from pathlib import Path
# import numpy as np
# from numpy import NaN
#
#
# def select_files():
#     return askopenfilenames(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])


# filenames = select_files()
# target_files = [file for file in filenames if file.endswith('.xlsx')]
#
# for i, file in enumerate(target_files, 1):
#     print(Path(file).stem)
# test = np.full(shape=7, fill_value=NaN)
# print(test)


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
    filename = askopenfilename(filetypes=[('CSV', '*.csv',)])
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

        self.df = pd.DataFrame()
        self.counter = 1

    def create_widgets(self):
        login_button = ttk.Button(self, text="Go", command=self.getvalues)
        login_button.grid(column=0, row=3, sticky=tk.NW, padx=5, pady=5)

        save_button = ttk.Button(self, text="save", command=self.save)
        save_button.grid(column=0, row=4, sticky=tk.NW, padx=5, pady=5)

    def getvalues(self):
        file = None
        file = select_file()
        _tmp = None
        _row = None
        _tmp = pd.read_csv(file, encoding="ISO-8859-1")
        _row = _tmp.iloc[9].values
        self.df[self.counter] = _row
        self.counter += 1

    def save(self):
        self.df.to_csv('merged.csv', index=False, header=False)

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

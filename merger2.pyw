import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfilenames, askopenfilename
from tkinter.filedialog import asksaveasfilename

import pandas as pd
import os
import urllib.parse
import numpy as np

from typing import Iterable


def flatten(items):
    """Yield items from any nested iterable; see Reference."""
    for x in items:
        if isinstance(x, Iterable) and not isinstance(x, (str, bytes)):
            for sub_x in flatten(x):
                yield sub_x
        else:
            yield x


def select_file():
    filename = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])
    return filename


def select_files():
    return askopenfilenames(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])


def extract_url(url, keys):
    res = urllib.parse.parse_qs(url)
    return pd.DataFrame.from_dict({key: value for (key, value) in res.items() if key in keys}, orient='columns')


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.geometry("500x300")
        self.title("Excel Appender")
        # self.resizable(0, 0)

        # configure the grid
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)

        # variables
        self.source_filename = None
        self.target_filename = None
        # self.source_dataframe = None
        # self.target_dataframe = None

        self.source_base_column = None
        self.target_base_column = None
        self.target_position = None
        self.url_keys = None

        # initialize dataframes
        self.source_dataframe = pd.DataFrame({})
        self.target_dataframe = pd.DataFrame({})

        # widgets
        self.source_entry = None
        self.source_excluded_columns = None
        self.target_entry = None
        self.target_excluded_columns = None
        self.url_source = None

        self.target_files = None

        # Initialize and create the UX widgets
        self.create_widgets()

    def create_widgets(self):
        # source
        # row 0
        ttk.Label(self, text="Base file:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

        self.source_entry = ttk.Entry(self)
        self.source_entry.grid(column=1, row=0, sticky=tk.EW, padx=5, columnspan=2)

        ttk.Button(self, text="Select", command=self.select_source).grid(column=3, row=0, sticky=tk.E, padx=5, pady=5)

        # row 1
        ttk.Label(self, text="Excluded columns: ").grid(column=0, row=1, sticky=tk.W, padx=5, pady=6)
        self.source_excluded_columns = Text(self, height=2, width=20)
        self.source_excluded_columns.grid(row=1, column=1, sticky=tk.EW, padx=5, columnspan=2)
        self.source_excluded_columns.insert(END, 'Device,Browser,IP,Start,End,Duration,Current Link,Route Part')

        # row 2
        ttk.Label(self, text="Base column: ").grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
        # self.source_base_column = ttk.Entry(self)
        # self.source_base_column.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)
        # self.source_base_column.insert(END, "Respondent ID")
        # menu = StringVar()
        # menu.set("Select base column")
        # self.source_base_column = ttk.OptionMenu(self, menu, [])

        # row 3
        ttk.Separator(self, orient="horizontal").grid(row=3, column=1, columnspan=2, sticky=tk.EW)

        # target label
        # row 4
        ttk.Label(self, text="Target file/s:").grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)

        self.target_entry = ttk.Entry(self)
        self.target_entry.grid(column=1, row=4, sticky=tk.EW, padx=5, pady=5, columnspan=2)

        ttk.Button(self, text="Select", command=self.select_target).grid(column=3, row=4, sticky=tk.E, padx=5, pady=5)

        # row 5
        ttk.Label(self, text="Excluded columns").grid(column=0, row=5, sticky=tk.W, padx=5, pady=6)
        self.target_excluded_columns = Text(self, height=2, width=20)
        self.target_excluded_columns.grid(row=5, column=1, sticky=tk.EW, padx=5, columnspan=2)
        self.target_excluded_columns.insert(END, 'Start,End,Duration,Current Link,Route Part')

        # row 6
        tk.Label(self, text="Base column: ").grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
        # self.target_base_column = ttk.Entry(self)
        # self.target_base_column.grid(row=6, column=1, sticky=tk.EW, pady=5, padx=6, columnspan=2)
        # self.target_base_column.insert(END, "Respondent ID")

        # row 7
        # tk.Label(self, text="Insert to index: ").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
        # self.target_position = ttk.Entry(self)
        # self.target_position.grid(row=7, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)
        # self.target_position.insert(END, "8")

        # # row 8
        # tk.Label(self, text="Extract data from URL: ").grid(row=8, column=0, sticky=tk.W, padx=5, pady=5)
        # self.url_keys = ttk.Entry(self)
        # self.url_keys.grid(row=8, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)
        # self.url_keys.insert(END, "URL:")
        #
        # # row 9
        # tk.Label(self, text="URL Source: ").grid(row=9, column=0, sticky=tk.W, padx=5, pady=5)
        # self.url_source = tk.StringVar(self)
        # dd = ttk.OptionMenu(self, self.url_source, "Source", "Source", "Target")
        # dd.grid(row=9, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)

        # row 10
        login_button = ttk.Button(self, text="Process", command=self.run_append2)
        login_button.grid(column=3, row=10, sticky=tk.E, padx=5, pady=5)

    def select_source(self):
        this = self
        self.source_filename = select_file()
        self.source_entry.delete(0, END)
        self.source_entry.insert(0, os.path.basename(self.source_filename))
        if this.source_filename:
            if this.source_filename.endswith('.csv'):
                this.source_dataframe = pd.read_csv(this.source_filename, encoding="ISO-8859-1")
            else:
                this.source_dataframe = pd.read_excel(this.source_filename)

        columns = this.source_dataframe.columns.values
        this.source_base_column = tk.StringVar(this)
        dd = ttk.OptionMenu(this, this.source_base_column, columns[0], *columns)
        dd.grid(row=2, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)

    def select_target(self):
        this = self
        filenames = select_files()
        self.target_files = [file for file in filenames if file.endswith('.xlsx')]
        first_file = filenames[0]
        # for i, file in enumerate(files[1:], 1):
        #     print(file)
        # this.target_filename = select_files()
        this.target_entry.delete(0, END)
        this.target_entry.insert(0, os.path.basename(first_file))

        if first_file.endswith('.csv'):
            df = pd.read_csv(first_file, encoding="ISO-8859-1")
        else:
            df = pd.read_excel(first_file)
        columns = df.columns.values
        this.target_base_column = tk.StringVar(this)
        dd = ttk.OptionMenu(this, this.target_base_column, columns[0], *columns)
        dd.grid(row=6, column=1, sticky=tk.EW, pady=5, padx=6, columnspan=2)

        this.target_position.delete(0, END)
        this.target_position.insert(0, len(columns))

    def run_append2(self):
        this = self
        header_collection = []
        source_excluded = self.source_excluded_columns.get("1.0", "end-1c").split(",")
        source_df = this.source_dataframe.loc[:, ~this.source_dataframe.columns.isin(source_excluded)]

        # get the column names from dataframes both the source and the target
        source_base_column = self.source_base_column.get()
        target_base_column = self.target_base_column.get()

        source_columns = source_df.columns
        header_collection.append(source_columns)
        t = {v: k for k, v in enumerate(source_columns) if v == source_base_column}
        source_base = t.get(source_base_column)
        source_df.columns = np.arange(0, len(source_columns))

        result = source_df[source_df.columns]
        files = self.target_files
        for i, file in enumerate(files, 1):
            count = len(result.columns)
            target_df = pd.read_excel(file)
            # get the excluded columns from the target file, then exclude the column from the dataframe
            target_excluded = self.target_excluded_columns.get("1.0", "end-1c").split(",")
            target_df = target_df.loc[:, ~target_df.columns.isin(target_excluded)]

            target_columns = target_df.columns

            header_collection.append(target_columns)
            t = {v: k for k, v in enumerate(target_columns) if v == target_base_column}
            target_base = t.get(target_base_column) + count
            target_df.columns = np.arange(count, len(target_columns) + count)

            result = result.merge(target_df[target_df.columns], left_on=source_base, right_on=target_base,
                                  how='left', suffixes=('_x', '_y'))

        headers = []
        headers.insert(0, list(flatten(header_collection)))

        result = pd.concat([pd.DataFrame(headers), result], ignore_index=True)
        # print(header_collection)
        filename = asksaveasfilename(filetypes=(("Excel", ('*.xls', '*.xlsx')), ("All files", '*.*')),
                                     defaultextension=".xlsx")
        result.to_excel(filename, index=False, header=False)


if __name__ == "__main__":
    app = App()
    app.mainloop()

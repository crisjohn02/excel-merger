import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
import pandas as pd
import os
import urllib.parse
import numpy as np


def select_file():
    filename = askopenfilename(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])
    return filename


def extract_url(url, keys):
    res = urllib.parse.parse_qs(url)
    return pd.DataFrame.from_dict({key: value for (key, value) in res.items() if key in keys}, orient='columns')


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

        # Initialize and create the UX widgets
        self.create_widgets()

    def create_widgets(self):
        # source
        # row 0
        ttk.Label(self, text="Source file:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

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
        ttk.Label(self, text="Target file:").grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)

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
        tk.Label(self, text="Insert to index: ").grid(row=7, column=0, sticky=tk.W, padx=5, pady=5)
        self.target_position = ttk.Entry(self)
        self.target_position.grid(row=7, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)
        self.target_position.insert(END, "8")

        # row 8
        tk.Label(self, text="Extract data from URL: ").grid(row=8, column=0, sticky=tk.W, padx=5, pady=5)
        self.url_keys = ttk.Entry(self)
        self.url_keys.grid(row=8, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)
        self.url_keys.insert(END, "URL:")

        # row 9
        tk.Label(self, text="URL Source: ").grid(row=9, column=0, sticky=tk.W, padx=5, pady=5)
        self.url_source = tk.StringVar(self)
        dd = ttk.OptionMenu(self, self.url_source, "Source", "Source", "Target")
        dd.grid(row=9, column=1, sticky=tk.EW, padx=5, pady=5, columnspan=2)

        # row 10
        login_button = ttk.Button(self, text="Process", command=self.run_append)
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
        this.target_filename = select_file()
        this.target_entry.delete(0, END)
        this.target_entry.insert(0, os.path.basename(this.target_filename))

        if this.target_filename.endswith('.csv'):
            this.target_dataframe = pd.read_csv(this.target_filename, encoding="ISO-8859-1")
        else:
            this.target_dataframe = pd.read_excel(this.target_filename)
        columns = this.target_dataframe.columns.values
        this.target_base_column = tk.StringVar(this)
        dd = ttk.OptionMenu(this, this.target_base_column, columns[0], *columns)
        dd.grid(row=6, column=1, sticky=tk.EW, pady=5, padx=6, columnspan=2)

        this.target_position.delete(0, END)
        this.target_position.insert(0, len(columns))

    def run_append(self):
        this = self
        lbl = tk.Label(self, text="Processing")
        lbl.grid(row=11, column=2, columnspan=2, padx=5, pady=5)

        # get the excluded columns from the source file, then exclude the column from the dataframe
        source_excluded = self.source_excluded_columns.get("1.0", "end-1c").split(",")
        source_df = this.source_dataframe.loc[:, ~this.source_dataframe.columns.isin(source_excluded)]

        # get the excluded columns from the target file, then exclude the column from the dataframe
        target_excluded = self.target_excluded_columns.get("1.0", "end-1c").split(",")
        target_df = this.target_dataframe.loc[:, ~this.target_dataframe.columns.isin(target_excluded)]

        self.source_dataframe = source_df.copy()
        self.target_dataframe = target_df.copy()

        # get the column names from dataframes both the source and the target
        source_base_column = self.source_base_column.get()
        target_base_column = self.target_base_column.get()

        to_be_appended = pd.DataFrame({})
        to_be_appended_url = pd.DataFrame({})

        # get URL extraction settings
        tmp = self.url_keys.get().split(":")
        url_base_column = tmp[0]
        url_columns_tolist = tmp[1].split(",") if len(tmp[1]) > 0 else None
        new_target = self.target_dataframe.copy()
        new_target = new_target[0:0]

        for index, row in self.target_dataframe.iterrows():
            new_target = pd.concat([new_target, pd.DataFrame([row.values], columns=row.keys())], ignore_index=True)
            # Extract data from source dataframe if there were any
            if self.source_filename:
                d = self.source_dataframe.loc[self.source_dataframe[source_base_column].isin([row[target_base_column]])]
                # d = d.loc[:, ~d.columns.isin([source_base_column])]
                if d.empty:
                    empty_array = np.full(len(self.source_dataframe.columns), None)
                    d = pd.DataFrame([empty_array], columns=self.source_dataframe.columns.values.tolist())
                if len(d.index) > 1:
                    count = len(d.index) - 1
                    while count > 0:
                        new_target = pd.concat([new_target, pd.DataFrame([row.values], columns=row.keys())],
                                               ignore_index=True)
                        count = count - 1
                to_be_appended = pd.concat([to_be_appended, d], ignore_index=True)

            # Extract URL
            if url_columns_tolist:
                _source = self.url_source.get()
                _sourceURL = None
                if _source == "Source":
                    _sourceURL = self.source_dataframe[url_base_column][index]
                else:
                    _sourceURL = row[url_base_column]

                if _sourceURL is not None:
                    extracted_url = extract_url(_sourceURL, url_columns_tolist)
                    to_be_appended_url = pd.concat([to_be_appended_url, extracted_url], ignore_index=True)

        position = len(self.target_dataframe.columns)
        to_be_appended_columns = []
        if self.source_filename:
            to_be_appended_columns = to_be_appended.columns.values.tolist()
            appended = {value: to_be_appended[value] for(value) in to_be_appended_columns}
            hj = pd.DataFrame(appended)
            or_columns = new_target.columns.values
            new_target = pd.concat([new_target, hj], ignore_index=True, axis=1)
            new_target.columns = np.concatenate((np.array(or_columns), np.array(to_be_appended_columns)))
            # for index, row in enumerate(to_be_appended_columns):
            #     if int(self.target_position.get()) != 0:
            #         position = int(self.target_position.get())
            #
            #     position += index
            #     new_target.insert(position, row, to_be_appended[row])

        if url_columns_tolist:
            position = int(self.target_position.get()) + len(to_be_appended_columns) - 1
            to_be_appended_url_columns = to_be_appended_url.columns.values.tolist()
            for idx, rw in enumerate(to_be_appended_url_columns):
                position += 1
                new_target.insert(position, rw, to_be_appended_url[rw])

        # Get the destination file name and save dataframe to a file with CSV format
        filename = asksaveasfilename(filetypes=(("Excel", ('*.xls', '*.xlsx')), ("All files", '*.*')), defaultextension=".xlsx")
        new_target.to_excel(filename, index=False, encoding='utf_8_sig')
        lbl.config(text="Data successfully appended!")


if __name__ == "__main__":
    app = App()
    app.mainloop()

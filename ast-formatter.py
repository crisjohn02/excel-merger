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

        login_button = ttk.Button(self, text="Align AST Tabular Excel", command=self.run_merge2)
        login_button.grid(column=0, row=3, sticky=tk.NW, padx=5, pady=5)

    def select_source(self):
        self.source_filename = select_file()
        self.source_entry.delete(0, END)
        self.source_entry.insert(0, os.path.basename(self.source_filename))

    def select_target(self):
        self.target_filename = select_file()
        self.target_entry.delete(0, END)
        self.target_entry.insert(0, os.path.basename(self.target_filename))

    def run_merge(self):
        df = pd.DataFrame({})
        excel_list = []
        files = glob.glob(self.target_filename + "/*")
        for file in files:

            excel_list.append(pd.read_excel(file))

        for ex in excel_list:
            df = df.append(ex, ignore_index=True)
        filename = asksaveasfilename(filetypes=(("Excel", '*.xlsx'), ("All files", '*.*')), defaultextension=".xlsx")
        print(filename)
        df.to_excel(filename, index=False)

    def run_merge2(self):
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
            base1 = base1.append(pd.read_excel(file, sheet_name="Computed Scores"))
            base2 = base2.append(pd.read_excel(file, sheet_name="IQ Scores"))
            base3 = base3.append(pd.read_excel(file, sheet_name="Counts"))

            base1.to_excel(writer, sheet_name="Computed Scores", index=False)
            base2.to_excel(writer, sheet_name="IQ Scores", index=False)
            base3.to_excel(writer, sheet_name="Counts", index=False)
            writer.save()

    def run_append(self):
        lbl = tk.Label(self, text="Processing")
        lbl.grid(row=11, column=2, columnspan=2, padx=5, pady=5)

        source_df = pd.DataFrame({})

        # Extract data from the source file and load to a panda's dataframe
        if self.source_filename:
            if self.source_filename.endswith('.csv'):
                source_df = pd.read_csv(self.source_filename, encoding="ISO-8859-1")
            else:
                source_df = pd.read_excel(self.source_filename)

        # Extract data from the target file and load to a panda's dataframe
        if self.target_filename.endswith('.csv'):
            target_df = pd.read_csv(self.target_filename, encoding="ISO-8859-1")
        else:
            target_df = pd.read_excel(self.target_filename)

        # get the excluded columns from the source file, then exclude the column from the dataframe
        source_excluded = self.source_excluded_columns.get("1.0", "end-1c").split(",")
        source_df = source_df.loc[:, ~source_df.columns.isin(source_excluded)]

        # get the excluded columns from the target file, then exclude the column from the dataframe
        target_excluded = self.target_excluded_columns.get("1.0", "end-1c").split(",")
        target_df = target_df.loc[:, ~target_df.columns.isin(target_excluded)]

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
        url_columns_tolist = tmp[1].split(",")

        for index, row in self.target_dataframe.iterrows():

            # Extract data from source dataframe if there were any
            if self.source_filename:
                d = self.source_dataframe.loc[self.source_dataframe[source_base_column].isin([row[target_base_column]])]
                d = d.loc[:, ~d.columns.isin([source_base_column])]
                to_be_appended = to_be_appended.append(d, ignore_index=True)

            # Extract URL
            if url_columns_tolist:
                extracted_url = extract_url(row[url_base_column], url_columns_tolist)
                to_be_appended_url = to_be_appended_url.append(extracted_url, ignore_index=True)

        position = len(self.target_dataframe.columns)
        to_be_appended_columns = []
        if self.source_filename:
            to_be_appended_columns = to_be_appended.columns.values.tolist()
            for index, row in enumerate(to_be_appended_columns):
                if int(self.target_position.get()) != 0:
                    position = int(self.target_position.get())

                position += index
                self.target_dataframe.insert(position, row, to_be_appended[row])

        if url_columns_tolist:
            position = int(self.target_position.get()) + len(to_be_appended_columns) - 1
            to_be_appended_url_columns = to_be_appended_url.columns.values.tolist()
            for idx, rw in enumerate(to_be_appended_url_columns):
                position += 1
                self.target_dataframe.insert(position, rw, to_be_appended_url[rw])

        # Get the destination file name and save dataframe to a file with CSV format
        filename = asksaveasfilename(filetypes=(("CSV", '*.csv'), ("All files", '*.*')), defaultextension=".csv")
        self.target_dataframe.to_csv(filename, index=False)
        lbl.config(text="Data successfully appended!")


if __name__ == "__main__":
    app = App()
    app.mainloop()

import glob
# import os.path

import pandas as pd
# import jpype
# import asposecells

# from tkinter import Tcl


# jpype.startJVM()
# from asposecells.api import Workbook

from natsort import natsorted


folder = "D:/Codes/python/py-apps/test"

# for filename in os.listdir(folder):
#     infilename = os.path.join(folder,filename)
#     if not os.path.isfile(infilename): continue
#     oldbase = os.path.splitext(filename)
#     newname = infilename.replace('.xls', '.html')
#     output = os.rename(infilename, newname)

# file_list = glob.glob(folder + "/*.html")
excl_list = []
# for file in file_list:
#     workbook = Workbook(file)
#     workbook.save(os.path.splitext(file)[0] + ".xlsx")
#     os.remove(file)
#
# jpype.shutdownJVM()


def re(text):
    return None if 'Unnamed' in text else text


outputxlsx = pd.DataFrame()
file_list2 = natsorted(glob.glob(folder + "/*.xlsx"))
print(file_list2)
for file in file_list2:
    df = pd.read_excel(file)
    header = []
    header.insert(0, list(map(re, df.columns.values)))
    # print()
    z = pd.concat([pd.DataFrame(header), df.transpose().reset_index(drop=True).transpose()], ignore_index=True)
    excl_list.append(z)
    excl_list.append(pd.DataFrame(pd.Series([None])))


    # excl_list = pd.concat([excl_list, pd.read_excel(file, sheet_name=None)], ignore_index=True, sort=False)
    # df.append(pd.Series(), ignore_index=True)

    # outputxlsx = outputxlsx.append(df, ignore_index=True)
#




# print(excl_list)
# concatenate all DataFrames in the list
# into a single DataFrame, returns new
# DataFrame.
# print(excl_list)
excl_merged = pd.concat(excl_list, ignore_index=True)

# exports the dataframe into excel file
# with specified name.
# writer = pd.ExcelWriter("dataframe.xlsx", engine='xlsxwriter')
excl_merged.to_excel('test.xlsx', index=False)
# writer.save()

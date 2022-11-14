from tkinter.filedialog import askopenfilenames
import os
from pathlib import Path
import numpy as np
from numpy import NaN


def select_files():
    return askopenfilenames(filetypes=[('Excel', ('*.xls', '*.xslm', '*.xlsx')), ('CSV', '*.csv',)])


# filenames = select_files()
# target_files = [file for file in filenames if file.endswith('.xlsx')]
#
# for i, file in enumerate(target_files, 1):
#     print(Path(file).stem)
test = np.full(shape=7, fill_value=NaN)
print(test)

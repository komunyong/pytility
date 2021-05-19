"""
To convert excel file into fixed-width file output (.txt)
Change 'Setting' section as you want. For styling text, please see how to use string format in detail at reference link

Reference: https://docs.python.org/3/library/string.html#format-specification-mini-language
{
  0: refer to first variable
  :: after colon symbol are all format mode
  <chars>: character to fills
  <^>: alignment [left, center, right]
  <number>: number of character to fills
}
"""

import pandas as pd
import numpy as np

# Setting
# TODO: Make Setting to be more dynamic, get input from command line typing with args()
# TODO: Make STYLES format to read array from JSON file to avoid to edit inside python script
INPUT_FILENAME = 'input.xlsx'
INPUT_SHEETNAME = 1  # or index of excel sheet
OUTPUT_FILENAME = 'output.txt'
STYLES = [
    '{0:0<10}', 
    '{0:<8}', 
    '{0:<10}', 
    '{0:<2}', 
    '{0:<15}', 
    '{0:0>10}', 
    '{0:<20}', 
    '{0:<20}', 
    '{0:<1}', 
    '{0:<1}', 
]

# Start program
if __name__ == '__main__':
    df = pd.read_excel(INPUT_FILENAME, sheet_name=INPUT_SHEETNAME)
    with open(OUTPUT_FILENAME, 'w') as f:
        for rows in df.values:
            print(rows)
            for i, val in enumerate(rows):
                f.write(STYLES[i].format(str(val)))
            f.write('\n')

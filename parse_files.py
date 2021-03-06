from pyxlsb import open_workbook, convert_date
import csv

""" 
filename to read from. can be setup to read multiple files.
"""
f = "Solar+Battery Example v12d (PPA+Bat).xlsb"
# f = "testfile.xlsb"
# output_file = "combined"

"""
filters the tabs to read by tab name, i.e. "PaR Inputs"
"""
tabsFilter = "inputs"
sheets = [s for s in open_workbook(f).sheets if s.lower().find(tabsFilter) != -1]

with open_workbook(f) as wb:
    for sheetname in sheets:
        with wb.get_sheet(sheetname) as sheet:
            # creates a csv for each tab in notebook with that tab's data
            with open(sheetname + ".csv", "w") as f:
                writer = csv.writer(f)
                for row in sheet.rows():
                    r = []
                    for c in row:
                        cell_value = c.v
                        # date cells get converted to number
                        # select date column to perserve format
                        # A = 0, B = 1, etc
                        if c.c == 1:
                            cell_value = convert_date(c.v)
                        r.append(cell_value)
                    writer.writerow(r)

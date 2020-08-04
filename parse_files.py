from pyxlsb import open_workbook, convert_date
import csv
from pathlib import Path
from datetime import datetime

now = datetime.now()
date_time = now.strftime("%m%d%Y_%H%M%S")
"""
name of file to write all the data
"""
output_file = "combined_" + date_time

"""
filters the tabs to read by tab name, i.e. "PaR Inputs"
"""
tabsFilter = "Results Summary"


def findFiles():
    p = Path(".")
    return Path(p).glob("**/*.xlsb")


def readFiles(paths):
    with open(output_file + ".csv", "a") as f:
        filecount = 0
        writer = csv.writer(f)
        for path in paths:
            d = str(path)
            # with open(output_file + ".csv", "a") as f:
            #     writer = csv.writer(f)
            with open_workbook(d) as wb:
                try:
                    with wb.get_sheet(tabsFilter) as sheet:
                        rowcount = 0
                        for row in sheet.rows():
                            rowcount += 1
                            if rowcount > 4:
                                # skip headers except for first file
                                if rowcount == 5 and filecount != 0:
                                    continue
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
                                continue
                    filecount += 1
                except ValueError as e:
                    pass
            wb.close()
    f.close()
    print(f"files processed: {filecount}. file created: {output_file}.csv")


paths = findFiles()
readFiles(paths)

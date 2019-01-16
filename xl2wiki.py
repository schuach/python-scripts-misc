#! /usr/bin/env/ python3
import sys
import xlrd
import pyperclip

if __name__ == '__main__':

    if not len(sys.argv) > 1:
        sys.exit("Usage: xl2wiki infile.xls [outfile]")

    # Read it in and build up a string in wiki format
    book = xlrd.open_workbook(sys.argv[1])
    s = ""

    # Loop over all sheets
    for sheet in book.sheets():
        s += f"Tabelle: {sheet.name}\n"
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                formstring = str(sheet.cell(row_index, col_index).value).replace("\n", " ").replace("|", "&#124;").replace("^", "&#94;")
                if row_index == 0:
                    s += "| *{}* ".format(formstring)
                else:
                    s += "| {}".format(formstring)
            s += '|\n'
        s += "\n\n"
    # Save it or print it
    if len(sys.argv) == 3:
        with open(sys.argv[2], 'w') as f:
            f.write(s)
    else:
        pyperclip.copy(s)
        print(s)

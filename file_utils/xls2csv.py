import xlrd
import csv
import glob
import csv
import os
import sys

def spinning_cursor():
    while True:
        for cursor in '|/-\\':
            yield cursor

spinner = spinning_cursor()

for xlsfile in glob.glob(os.path.join('.', '*.xlsx')):

    wb = xlrd.open_workbook(xlsfile)
    sh = wb.sheet_by_name('Sheet1')
    filename = xlsfile + '.csv'
    print "parsing: " + filename
    new_csv_file = open(filename, 'wb')
    
    wr = csv.writer(new_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        sys.stdout.write(spinner.next())
        sys.stdout.flush()
        wr.writerow(sh.row_values(rownum))
        sys.stdout.write('\b')

    new_csv_file.close()


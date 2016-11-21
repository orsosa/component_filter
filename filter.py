#!/usr/bin/env python
import xlrd
from sys import argv

words=[]
col=5

def get_options(argv):
    global words
    global col
    for i in range(len(argv)):
        if argv[i] in ["--words","-w"]:
            try:
                words=argv[i+1].lower().split()
            except:
                print "you must supply a list of key words to look for after option %s"%argv[i]
                print 'e.g. ./filter.py -w "Comparator, comparador"'

        if argv[i] in ["--column","-c"]:
            try:
                col=int(argv[i+1])
            except:
                print "you must supply a column number after option %s"%argv[i]
                print 'e.g. ./filter.py -w "Comparator, comparador"'




if __name__=="__main__":
    xl_workbook = xlrd.open_workbook("Lab_Microelectronica_Noviembre_2016.xlsx")
    sheet_names = xl_workbook.sheet_names()
    xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])
    get_options(argv)
    print "words to filter: %s"%words
    print "printing column: %d"%col
    num_rows =xl_sheet.nrows
    for row_idx in range(num_rows):
        cell_obj = xl_sheet.cell(row_idx, 2)
        test=False
        for i in range(len(words)):
            test=test or (words[i] in cell_obj.value.lower())
#        print cell_obj.value.lower()
        if (test):
            print ("%s"%xl_sheet.cell(row_idx, col).value)

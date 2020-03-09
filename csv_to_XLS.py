'''


'''

import sys
import os
import csv
import struct
import datetime
import re
import openpyxl


def main():
    
    inFile = os.path.abspath(sys.argv[1])
    inFilename = os.path.basename(inFile)
    outFilename = ".".join(inFilename.split(".")[:-1])
    outputDir = os.path.dirname(inFile)
    outExcel = os.path.join(outputDir, "{}.xlsx".format(outFilename))

    with open(inFile, 'rb') as dataIn:
        
        print "Parsing data"
        
        csvIn = csv.reader(dataIn, quoting=csv.QUOTE_ALL)


        # Initialize the Excel Workbook and worksheet
        wb = openpyxl.Workbook()
        ws = wb.create_sheet("Data", 0)

        
        # Add record rows to worksheet
        print "Adding records to worksheet"
        for line in csvIn:
            ws.append(line)  

        # Save Excel workbook
        print "Saving workbook"
        wb.save(outExcel)


if __name__ == '__main__':
    main()

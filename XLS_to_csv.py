'''
Script to save an Excel file and a CSV.

Script creates a text file containing a VB script.

The VB script is run on the excel file via Command Line 
using Windows Script Host ("CScript") and outputs the CSV.

--> cscript   temp_vb_script   excel_file   csv_file <--

After the CSV is created, the VB script file is deleted.

Finally, script strips end spaces of each field and adds quotes.
'''


import os
import sys
import csv
import subprocess

def main():
    
    # Excel file to be processed
    excel_file = os.path.abspath(sys.argv[1])
    
    # In and out locations for all files
    outdir = os.path.dirname(excel_file)

    # CSV file to write Excel output to.
    csv_file = createOutputCSV(outdir, excel_file)
    
    # Temp XLS to CSV vb script
    temp_vb_script = createTempXLStoCSV(outdir)
    
    # Process using command line arguments
    subprocess.call(["cscript", temp_vb_script, excel_file, csv_file])
    
    # Delete temp vb script
    os.remove(temp_vb_script)
    
    # Quote and strip extra spaces CSV file
    quoteAndStripCsv(csv_file)
   

def createOutputCSV(outdir, excel_file):
    """ CSV file to write Excel output to. """
    
    excel_name = os.path.basename(excel_file)
    csv_name = ".".join(excel_name.split(".")[:-1]) + ".csv"
    csv_file = os.path.join(outdir, csv_name)
    return csv_file
    

def createTempXLStoCSV(outdir):
    """ Create temp VB script which will create the 
    CSV. Can't use tempfile since 'cscript' looks for 
    the file extension of the script.  """
    
    temp_vb_script = os.path.join(outdir, "tempXLStoCSV-DO_NOT_TOUCH.vbs")
    
    vb_string = "\n".join(
    
    ["Dim oExcel",
     "Set oExcel = CreateObject(\"Excel.Application\")",
     "oExcel.DisplayAlerts = False",
     "Dim oBook",
     "Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))",
     "oBook.SaveAs WScript.Arguments.Item(1), 6",
     "oBook.Close False",
     "oExcel.Quit",
     "WScript.Echo \"Done\""
    ])

    with open(temp_vb_script, 'wb') as t:
        t.write(vb_string)
    
    return temp_vb_script

    
def quoteAndStripCsv(csv_file):
    """ Read the contents of the CSV. Trim extra 
    spaces and add quotes to the fields. """
    
    contentList = []
    
    with open(csv_file, 'rb') as In:
        csvIn = csv.reader(In, quoting=csv.QUOTE_ALL)
        
        for row in csvIn:
            line = [' '.join(field.split()) for field in row]
            contentList.append(line)
            
    with open(csv_file, 'wb') as Out:
        csvOut = csv.writer(Out, quoting=csv.QUOTE_ALL)
            
        for record in contentList:
            csvOut.writerow(record)
   
        
if __name__ == '__main__':
    main()

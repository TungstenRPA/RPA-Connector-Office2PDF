##############################################################
############### Excel to PDF converter script ################
##############################################################
#
# Requirements:
#   - Windows OS (tested with Windows 10)
#   - Microsoft Excel (tested with Office 2013)
#   - Python 3.5 or newer
#   - comtypes https://pypi.org/project/comtypes/
#
# Author: Robert Birkenheuer
# Version: 0.1
#
##############################################################

import sys
import os
import comtypes.client

def excel2pdf(excelfile, pdffile, overwrite = 0):
    """
    Converts an Excel document to PDF

    Parameters:
        excelfile (string): absolute path and file name of the source excel document
        pdffile (string): absolute path and file name of the target PDF document
        overwrite (int): overwrite exitsting files, 1 = true, 0 = false

    Returns:
        string: ok for success or error message
    """

    xlTypePDF = 0
    xlQualityMinimum = 1
    xlOpenAfterPublish = 0
    response_text = ""

    if os.path.isfile(excelfile) == False:
        response_text = "File does not exist: " + excelfile
    elif overwrite == 0 and os.path.isfile(excelfile):
        response_text = "File does already exist: " + pdffile
    else:

        try:
            excel = comtypes.client.CreateObject('Excel.Application')
            try:
                excel.Visible = False
                doc = excel.Workbooks.Open(excelfile)
                try:
                    # doc.SaveAs(pdffile, FileFormat=wdFormatPDF)
                    doc.ExportAsFixedFormat(xlTypePDF, pdffile, xlQualityMinimum, xlOpenAfterPublish)
                    response_text = "ok"
                except:
                    response_text = "Failed to save file: " + pdffile # sys.exc_info()
                finally:
                    doc.Close()
            except:
                response_text = "Failed to open file: " + excelfile # sys.exc_info()
            finally:
                excel.Quit()
        except:
            response_text = "Failed to create Word Application" # sys.exc_info()

    return response_text

# For testing purpose if the script is called via shell
if __name__ == '__main__':
    in_file = os.path.abspath(sys.argv[1])
    out_file = os.path.abspath(sys.argv[2])
    over_write = os.path.abspath(sys.argv[3])
    print(excel2pdf(in_file, out_file, over_write))
##############################################################
############ Powerpoint to PDF converter script ##############
##############################################################
#
# Requirements:
#   - Windows OS (tested with Windows 10)
#   - Microsoft Powerpoint (tested with Office 2013)
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

def ppt2pdf(pptfile, pdffile, overwrite = 0):
    """
    Converts a PowerPoint pptx document to PDF

    Parameters:
        pptfile (string): absolute path and file name of the source PowerPoint document
        pdffile (string): absolute path and file name of the target PDF document
        overwrite (int): overwrite exitsting files, 1 = true, 0 = false

    Returns:
        string: ok for success or error message
    """

    pptFormatPDF = 32
    msoFalse = 0 # starts powerpoint without opening a window
    boolReadOnly = 0
    boolOpenUntitled = 0
    response_text = ""

    if os.path.isfile(pptfile) == False:
        response_text = "File does not exist: " + pptfile
    elif overwrite == 0 and os.path.isfile(pdffile):
        response_text = "File does already exist: " + pdffile
    else:

        try:
            powerpoint = comtypes.client.CreateObject('Powerpoint.Application')
            try:
                #powerpoint.Visible = 0
                slides = powerpoint.Presentations.Open(pptfile,boolReadOnly,boolOpenUntitled,msoFalse)
                try:
                    slides.SaveAs(pdffile, FileFormat=pptFormatPDF)
                    response_text = "ok"
                except:
                    response_text = "Failed to save file: " + pdffile # sys.exc_info()
                finally:
                    slides.Close()
            except:
                response_text = "Failed to open file: " + pptfile # sys.exc_info()
            finally:
                powerpoint.Quit()
        except:
            response_text = "Failed to create Powerpoint Application" # sys.exc_info()

    return response_text

# For testing purpose if the script is called via shell
if __name__ == '__main__':
    in_file = os.path.abspath(sys.argv[1])
    out_file = os.path.abspath(sys.argv[2])
    over_write = os.path.abspath(sys.argv[3])
    print(ppt2pdf(in_file, out_file, over_write))
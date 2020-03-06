##############################################################
############### Word to PDF converter script #################
##############################################################
#
# Requirements:
#   - Windows OS (tested with Windows 10)
#   - Microsoft Word (tested with Office 2013)
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

def word2pdf(wordfile, pdffile, overwrite = 0):
    """
    Converts a Word docx document to PDF

    Parameters:
        wordfile (string): absolute path and file name of the source word document
        pdffile (string): absolute path and file name of the target PDF document
        overwrite (int): overwrite exitsting files, 1 = true, 0 = false

    Returns:
        string: ok for success or error message
    """

    wdFormatPDF = 17
    response_text = ""

    if os.path.isfile(wordfile) == False:
        response_text = "File does not exist: " + wordfile
    elif overwrite == 0 and os.path.isfile(pdffile):
        response_text = "File does already exist: " + pdffile
    else:

        try:
            word = comtypes.client.CreateObject('Word.Application')
            try:
                word.Visible = False
                doc = word.Documents.Open(wordfile)
                try:
                    doc.SaveAs(pdffile, FileFormat=wdFormatPDF)
                    response_text = "ok"
                except:
                    response_text = "Failed to save file: " + pdffile # sys.exc_info()
                finally:
                    doc.Close()
            except:
                response_text = "Failed to open file: " + wordfile # sys.exc_info()
            finally:
                word.Quit()
        except:
            response_text = "Failed to create Word Application" # sys.exc_info()

    return response_text

# For testing purpose if the script is called via shell
if __name__ == '__main__':
    in_file = os.path.abspath(sys.argv[1])
    out_file = os.path.abspath(sys.argv[2])
    over_write = os.path.abspath(sys.argv[3])
    print(word2pdf(in_file, out_file, over_write))
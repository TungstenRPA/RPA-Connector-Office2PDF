##############################################################
################## Send Email with Outlook ###################
##############################################################
#
# Requirements:
#   - Windows OS (tested with Windows 10)
#   - Microsoft Outlook (tested with Office 2013)
#   - Python 3.5 or newer
#   - comtypes https://pypi.org/project/comtypes/
#
# Author: Robert Birkenheuer
# Version: 0.1
#
##############################################################

import sys
import os
from comtypes.client import CreateObject, GetActiveObject 
from outlook_instance import OutlookInstance

def send_mail(mailbox, to, subject, body, isHtml):
    """
    Send email

    Parameters:
        mailbox (string): name of outlook mailbox to use (e.g. some.name@company.com)
        to (string): receipient, multiple receipients separeted by ;
        subject (string): email subject
        body (string): email body
        isHtml (number): 0 = plain text; 1 = HTML email
    Returns:
        status (string): ok or error message
    """
    response_text = ""
    outlook = None
    try:
        outlook = OutlookInstance(mailbox)
        if isHtml:
            outlook.create_mail(to, subject, None, body, att_path=None, auto_send=True) 
        else:
            outlook.create_mail(to, subject, body, html_body=None, att_path=None, auto_send=True)
            
        response_text = "ok"
    except TypeError as typeerr:
        response_text = str(typeerr)
    except:
        response_text = str(sys.exc_info()[1])
    finally:
        if outlook and outlook.newinstance:
            outlook.close()

    return response_text

def send_mail2(mailbox, to, subject, body, isHtml, attachments):
    """
    Send email with attachments

    Parameters:
        mailbox (string): name of outlook mailbox to use (e.g. some.name@company.com)
        to (string): receipient, multiple receipients separeted by ;
        subject (string): email subject
        body (string): email body
        isHtml (number): 0 = plain text; 1 = HTML email
        to (string): attachment, multiple attachment separeted by ;
    Returns:
        status (string): ok or error message
    """
    response_text = ""
    outlook = None
    try:
        outlook = OutlookInstance(mailbox)
        if isHtml:
            outlook.create_mail(to, subject, None, body, attachments, auto_send=True) 
        else:
            outlook.create_mail(to, subject, body, None, attachments, auto_send=True)
            
        response_text = "ok"
    except TypeError as typeerr:
        response_text = str(typeerr)
    except:
        response_text = str(sys.exc_info()[1])
    finally:
        if outlook and outlook.newinstance:
            outlook.close()

    return response_text

    

# For testing purpose if the script is called via shell
if __name__ == '__main__':
    #in_file = os.path.abspath(sys.argv[1])
    #out_file = os.path.abspath(sys.argv[2])
    #print(send_mail("robert.birkenheuer@kofax.com", "kapow@gmx.de", "test subject", "<h1>test from robot as html</h1><p>with html</p>", 1))
    print(send_mail2("robert.birkenheuer@kofax.com", "kapow@gmx.de; robert.birkenheuer@kofax.com", "test subject", "<h1>test from robot as html</h1><p>with html</p>", 1, "c:/temp/Demo3.pdf;c:/temp/Demo4.pdf"))
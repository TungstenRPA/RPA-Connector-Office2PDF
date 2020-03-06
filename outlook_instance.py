##############################################################
################### Outlook COM Interface ####################
##############################################################
#
# Requiremenst:
#   - Windows OS (tested with Windows 10)
#   - Microsoft Outlook (tested with Office 2013)
#   - Python 3.5 or newer
#   - comtypes https://pypi.org/project/comtypes/
#
# Reference: https://docs.microsoft.com/en-us/office/vba/api/overview/outlook/object-model
#
##############################################################

import os

from comtypes.client import CreateObject, GetActiveObject    

class OutlookInstance():
    """Class to help with MS Outlook interoperability
    
    Attributes:
        outlook (comtypes.POINTER(_Application)): MS Outlook Application
        namespace (comtypes.POINTER(_NameSpace)): Namespace for outlook; Default = 'MAPI'
        mailbox (comtypes.POINTER()): Mailbox for target Outlook account
        targetfolder (comtypes.POINTER()): Outlook folder in mailbox to search
        delfolder (comtypes.POINTER()): Deleted Items folder for mailbox
        newinstance (boolen): True = new outlook instance was created; False = instance already existed
        
    Methods:
        set_namespace: Sets the namespace for outlook object
        set_mailbox: Sets the mailbox in namespace, if present.
        set_target_folder: Sets the targetfolder in mailbox
        create_mail: Create an email item to modify or send automatically
        close: dispose of the OutlookInstance object; close Outlook Application
    """
    
    class Oli():
        """Iterator for objects in the Outlook object.
            
        Attributes:
            _obj (outlook_object): object from an outlook application that contains items
            
        Methods:
            items: Yields the items for _obj
            prop: Gets properties for the items in _obj
        """
    
        def __init__(self, outlook_object):
            self._obj = outlook_object
            
        def items(self):
            array_size = self._obj.Count
            for item_index in range(1, array_size + 1):
                yield (item_index, self._obj[item_index])
                
        def prop(self):
            return sorted(self._obj._prop_map_get_.keys())
    
    def __enter__(self):
        return self
    
    def __init__(self, mailbox, namespace='MAPI', tgtfolder=None):
        """Initializes the object
        
        Args:
            mailbox (str): Name of the mailbox to use
            namespace (str): Namespace to use in Outlook; Default = 'MAPI'
            tgtfolder (str): Name of the folder in mailbox to search; Default = None
        """
        try:
            self.outlook = GetActiveObject('Outlook.Application')
            self.newinstance = False
        except:
            self.outlook = CreateObject('Outlook.Application')
            self.newinstance = True
        
        self.namespace = self.outlook.GetNamespace(namespace)
        self.set_mailbox(mailbox)
        
        if tgtfolder:
            self.set_target_folder(tgtfolder)
            
        #print('Outlook Instance initialized.')
            
    def set_namespace(self, namespace):
        """Sets the NameSpace for the Outlook Application
        
        Args:
            namespace (str): NameSpace to search for
        """
        self.namespace = self.outlook.GetNamespace(namespace)
        
    def set_mailbox(self, mailbox):
        """Sets the mailbox to use in the Outlook object
        
        Args:
            mailbox (str): Mailbox to search for
        """
        found = False
        for _, mbox in self.Oli(self.namespace.Folders).items():
            if mailbox.upper() in mbox.Name.upper():
                self.mailbox = mbox
                found = True
                break
        
        if not found:
            raise TypeError('Mailbox: {} not found!'.format(mailbox))
            #print('Mailbox: {} not found!'.format(mailbox))

    def set_target_folder(self, search_folder):
        """Sets the target folder in the mailbox
        
        Args:
            search_folder (str): Name of the folder to search for
            
        Methods:
            list_folders: recursively searches for the target folder, returns it when found
        """
        
        def list_folders(olObject,target):
            """Recursive function to search all folders in an Outlook object for a target folder
            
            Args:
                olObject (comtypes.POINTER()): Object containing folders to be searched
                target (str): Name of the folder to search for

            Returns:
                found_folder (comtypes.POINTER()): Folder with name matching target
            """
            nonlocal found_folder
            for _, folder in olObject.items():
                if found_folder is None:
                    if folder.Name.upper() == target.upper():
                        found_folder = folder
                    elif 'Deleted Items' in folder.Name:
                        self.delfolder = folder
                        list_folders(self.Oli(folder.Folders), target)
                    else:
                        list_folders(self.Oli(folder.Folders), target)
                else:
                    return found_folder
                
        found_folder = None
        list_folders(self.Oli(self.mailbox.Folders), search_folder)
        
        if found_folder is None:
            raise TypeError('Folder {} not found in mailbox {}!'.format(search_folder, self.mailbox.Name))
        else:
            self.targetfolder = found_folder
                             
    def create_mail(self, to=None, subject=None, body=None, html_body=None, att_path=None, auto_send=False):
        """Creates an Outlook MailItem
        
        Args:
            to (str): List of email addresses to add to the To: list, separated by ;
            subject (str): Subject line of the email
            body (str): Body of the email
            html_body (str): Body of the email in HTML
            att_path (str): Path to the files to attach, separated by ;
            auto_send (bool): Whether to send the email automatically or display the item to be modified further
        """
        olMailItem = 0x0
        
        item = self.outlook.CreateItem(olMailItem)
        item.To = to
        item.Subject = subject
        item.Body = body
        if html_body:
            item.HTMLBody = html_body

        if att_path:
            list_att = att_path.split(";")
            if isinstance(list_att, list):
                for attstr in list_att:
                    item.Attachments.Add(attstr)

        if auto_send:
            item.Send()
        else:
            item.Display()

    def close(self):
        self.namespace.Logoff()
        self.outlook.Quit()
        self.namespace = None
        self.outlook = None
        
    def __exit__(self, *args, **kwargs):
        self.close()

    
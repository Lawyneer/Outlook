#!/usr/bin/python3
"""
Last updated on June 21, 2021
@author: Lawyneer

"""

# Import Python packages/modules needed
import win32com.client as client

# Function to send emails
def email_func(to_emails, subject, msg, cc_emails, bcc_emails, display):
    """
    Parameters:
    ----------
    to_emails : str, list, or tuple
        The email address in which the email is to be sent.

    subject : str
        The subject of the email.

    msg : str
        The body of the message.

    cc_emails : str, list, or tuple
        The email addresses in which to copy the email.

    bcc_emails : str, list, or tuple
        The email addresses in which the email is to be sent via BCC.

    Returns:
    -------
    True or False : bool
            Boolean to specify if email was sent or not.

    Notes:
        See additional documentation see:
        https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem
                
    """

    # Extract emails from lists or tuples
    if type(to_emails) == list or type(to_emails) == tuple:
        seperator = '; '
        to_emails = seperator.join(to_emails)

    if type(cc_emails) == list or type(cc_emails) == tuple:
        seperator = '; '
        cc_emails = seperator.join(cc_emails)

    if type(bcc_emails) == list or type(bcc_emails) == tuple:
        seperator = '; '
        bcc_emails = seperator.join(bcc_emails)

    # Try to send the email
    try:
        outlook = client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = to_emails
        message.CC = cc_emails
        message.BCC = bcc_emails
        message.Subject = subject
        message.Body = msg
        if display == True:
            message.Display()
        else:
            message.Send()
        return True
    except:
        print('ERROR: No email sent.')
        return False
# # end submit function

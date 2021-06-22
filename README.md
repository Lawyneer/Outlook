# Outlook
A collection of functions to interact with Outlook using Python.  These are being written to help me learn about interacting with other programs using a program I write.

## email_func
This function takes email addreses, a subject, a message, and a boolean and sends am email to the email addresses using the supplied message and subject.
### Notes:
  1.  to_emails, cc_emails, and bcc_emails can be a string of email addresses seperated by a ";" or "," as well as a list or tuple of email addresses.
  2.  If nobody is to be copied or blind copied, enter enpty strings (i.e., `''`, or `""`) to the fourth or fifth positions.
  3.  The 'display' argument determines if the email will be sent without Outlook opening a window or if the email composition window will open.  If `position = False` then the email composition window opens and the user must hit "send" to actually send the email.  If `position = True` the email will be sent in the background without the user seeing it.  The sent email will be saved in the user "sent" folder in Outlook.

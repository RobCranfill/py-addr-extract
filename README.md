# py-addr-extract
Python program to extract email addresses from an online inbox, creating a CSV file suitable for import into Google Contacts.

! 
This program will use IMAP to read the specified inbox (note: other Inboxes?) from your online email provider.
It will look at all "To:", "From:", and "CC:" headers, and create a CSV file containing each distinct name,
along with up to 4 associated email addresses. (note: this is a shortcoming of the GMail import format: a max
of 4 addresses are allowed. Workaround?)

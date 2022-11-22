# py-addr-extract
Python program to extract email addresses from an online inbox and create a CSV file suitable for importing into Google Contacts.

## Overview
This program will use IMAP to read the specified inbox (note: other Inboxes?) from your online email provider.
It will look at all "To:", "From:", and "CC:" headers, and create a CSV file containing each distinct name,
along with up to 4 associated email addresses. (note: this is a shortcoming of the GMail import format: a max
of 4 addresses are allowed. Workaround?)

## Requirements
* Python 3.x
* Libraries (all standard libs)
  * sys
  * csv
  * email
  * imaplib

## Usage
There are 3 arguments to invoke this with:
* HOST Your email host name, like mail.gmail.com
* USER Your user name
* PASSWORD Your email password

 For example,

   `python3 extract.py imap.gmail.com joe.blow@gmail.com 16_char_app_pwd!`

## Limitations
* (MAJOR ISSUE) Does _not_ ignore case, so it will consider Foo@fum.com and foo@Fum.com to be different. A semi-easy
  fix to the code would be to use some sort of case-insensitive dictionary.
* Only associates a maximum of 4 email addresses with a name. This is a limitation of the Google Contacts CSV format,
  for which there is probably a work-around (create more than one entry and let Google Contacts 'merge' them?)

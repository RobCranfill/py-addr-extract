import imaplib
import email
import csv
import sys
"""
Extract headers from online email and create a CSV file
 to import all found email addresses as contacts into GMail.

(c)2022 robcranfill@gmail.com
based on https://qr.ae/pvQWFi

Command-line args: {HOST} {USERNAME} {PASSWORD}

"""

skipped = 0


def split_addrs(s):
    """
    split an address list into list of tuples of (name,address)
    """
    if not(s):
        return []
    outQ = True
    cut = -1
    res = []
    for i in range(len(s)):
        if s[i] == '"':
            outQ = not(outQ)
        if outQ and s[i] == ',':
            res.append(email.utils.parseaddr(s[cut+1:i]))
            cut = i
    res.append(email.utils.parseaddr(s[cut+1:i+1]))
    return res


def accumulate_goodness(biggus_dictus, list_of_name_and_address):
    """
    Biggus dictus! :-)
    """
    global skipped
    if len(list_of_name_and_address) == 0:
        return

    for p in list_of_name_and_address:
        user_name = p[0]
        email_addr = p[1].lower()
        # print(f" name: '{email_name}', addr: '{email_addr}'")

        if len(user_name) == 0:
            # print("   skipping empty name")
            skipped += 1
            return
        if user_name.startswith("=?"):  # common bullshit - FIX?
            skipped += 1
            return

        if user_name in biggus_dictus:
            addresses = biggus_dictus[user_name]
        else:
            print(f"  new name '{user_name}'")
            addresses = set()
            biggus_dictus[user_name] = addresses
        # print(f"  Prev set of addresses is {addresses}")
        addresses.add(email_addr)
        if len(addresses) == 4:
            print(f"  **** MAX ADDRESSES HIT FOR '{user_name}'")


# begin
if len(sys.argv) != 4:
    print(f"{sys.argv[0]} HOST USERNAME PASSWORD")
    exit(1)
pHost = sys.argv[1]
pUser = sys.argv[2]
pPassword = sys.argv[3]
print(f"Loging in with '{pHost}', '{pUser}', '{pPassword}'...")

mail = imaplib.IMAP4_SSL(pHost)
mail.login(pUser, pPassword)
mail.select("INBOX")
result, data = mail.search(None, "ALL")
ids = data[0].split()
# data_as_str = bytes(data).decode("utf-8")
# ids = data_as_str[0].split()

# FOR DEBUG
limit = 10000
print(f"SHORTENING DATA TO FIRST {limit} ITEMS....")
ids = ids[0:limit]

# PROBLEM: ids is a tuple of bytes, but needs to be strings! (or does it?)
# super cheesy way to fix:
ids_str = []
for id in ids:
    ids_str.append(id.decode('UTF-8'))

print(f"fetching {len(ids)} ids....")
msgs = mail.fetch(','.join(ids_str), '(BODY.PEEK[HEADER])')[1][0::2]

print(f"fetched {len(msgs)} messages; iterating....")

# intermediate result: dictionary of
#  key: Person (or company) name
#  value: set of email addresses associated with that person
#
names_and_addresses = {}

for x, msg in msgs:
    msg_str = msg.decode('UTF-8')
    # print(f"Raw message is\n----------\n{msg_str}\n----------\n")

    # was msg; TODO: could use message_from_bytes and avoid above conversion??
    msgobj = email.message_from_string(msg_str)
    # msgobj = email.message_from_bytes(msg)

    accumulate_goodness(names_and_addresses, split_addrs(msgobj['to']))
    accumulate_goodness(names_and_addresses, split_addrs(msgobj['from']))
    accumulate_goodness(names_and_addresses, split_addrs(msgobj['cc']))

print("------------------------------------------------")
print(f"Found {len(names_and_addresses)} distinct names")
print(f" skipped: {skipped}")
# print(f"{names_and_addresses}")

# Now create a list of dictionaries for the CSV DictWriter.
# Each entry in the list is a dictionary like this:
#  {"Name": "Rob Cranfill",
#   "E-mail 1 - Value": "robcranfill@comcast.net",
#   "E-mail 2 - Value": "robcranfill@gmail.com"}
#
csv_data = []
for name_key in names_and_addresses.keys():
    d = dict()
    d["Name"] = name_key
    i = 1
    # todo: check for email addresses that only differ in upper/lower case?
    # todo: if there are more than 4 addresses, will create bad CSV?
    for a in names_and_addresses[name_key]:
        if i > 4:
            print(
                f"Warning! Email address #{i} found for name {name_key}! Skipping!")
            continue
        e_key_name = f"E-mail {i} - Value"
        d[e_key_name] = a
        # print(f" >>> {key_name} = {a}")
        i += 1
    csv_data.append(d)

# print(f"\n\nCSV dict: {csv_data}")

# Write the CSV file!
with open('output.csv', 'w', newline='') as csvfile:
    Gfieldnames = ["Name", "Given Name", "Additional Name", "Family Name",
                   "Yomi Name", "Given Name Yomi", "Additional Name Yomi",
                   "Family Name Yomi", "Name Prefix", "Name Suffix",
                   "Initials", "Nickname", "Short Name", "Maiden Name",
                   "Birthday", "Gender", "Location", "Billing Information",
                   "Directory Server", "Mileage", "Occupation", "Hobby",
                   "Sensitivity", "Priority", "Subject", "Notes",
                   "Language", "Photo", "Group Membership",
                   "E-mail 1 - Type", "E-mail 1 - Value",
                   "E-mail 2 - Type", "E-mail 2 - Value",
                   "E-mail 3 - Type", "E-mail 3 - Value",
                   "E-mail 4 - Type", "E-mail 4 - Value",
                   "Phone 1 - Type", "Phone 1 - Value",
                   "Phone 2 - Type", "Phone 2 - Value",
                   "Phone 3 - Type", "Phone 3 - Value",
                   "Phone 4 - Type", "Phone 4 - Value",
                   "Address 1 - Type", "Address 1 - Formatted",
                   "Address 1 - Street", "Address 1 - City",
                   "Address 1 - PO Box", "Address 1 - Region",
                   "Address 1 - Postal Code", "Address 1 - Country",
                   "Address 1 - Extended Address",
                   "Address 2 - Type", "Address 2 - Formatted",
                   "Address 2 - Street", "Address 2 - City",
                   "Address 2 - PO Box", "Address 2 - Region",
                   "Address 2 - Postal Code", "Address 2 - Country",
                   "Address 2 - Extended Address", "Organization 1 - Type",
                   "Organization 1 - Name", "Organization 1 - Yomi Name",
                   "Organization 1 - Title", "Organization 1 - Department",
                   "Organization 1 - Symbol", "Organization 1 - Location",
                   "Organization 1 - Job Description", "Website 1 - Type",
                   "Website 1 - Value"]

    writer = csv.DictWriter(csvfile, fieldnames=Gfieldnames)
    writer.writeheader()
    for d in csv_data:
        writer.writerow(d)

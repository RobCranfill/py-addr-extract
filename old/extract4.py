import imaplib
import email
import csv
import sys
# extract headers from online email! works!
# robcranfill@gmail.com
# based on https://qr.ae/pvQWFi


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


# begin

if len(sys.argv) != 4:
    print(f"{sys.argv[0]} HOST USERNAME PASSWORD")
    exit(1)
pAddress = sys.argv[1]
pUser = sys.argv[2]
pPassword = sys.argv[3]

mail = imaplib.IMAP4_SSL(pAddress)
print("mail connection established....")

mail.login(pUser, pPassword)
print("logged in....")

mail.select("INBOX")
print("selected mailbox....")

result, data = mail.search(None, "ALL")
print("search OK....")


ids = data[0].split()
# data_as_str = bytes(data).decode("utf-8")
# ids = data_as_str[0].split()

# FOR DEBUG only
ids = ids[0:100]
print("SHORTENING DATA....")

# PROBLEM: ids is a tuple of bytes, but needs to be strings! (or does it?)
# super cheesy way to fix:
ids_str = []
for id in ids:
    ids_str.append(id.decode('UTF-8'))
# print(f"id as str is {ids_str}")

print(f"fetching {len(ids)} ids....")
# was:
#   msgs = mail.fetch(','.join(ids), '(BODY.PEEK[HEADER])')[1][0::2]
# no go:
#   msgs = mail.fetch(','.join([str(id) for value in ids]), '(BODY.PEEK[HEADER])')[1][0::2]
msgs = mail.fetch(','.join(ids_str), '(BODY.PEEK[HEADER])')[1][0::2]

print(f"fetched {len(msgs)} messages; iterating....")

# result: dictionary of
#  key: Person (or company) name
#  value: set of email addresses associated with that person
#
names_and_addresses = {}

iter = 0
skipped = 0
for x, msg in msgs:
    iter += 1
    # print(f"parsing #{iter}....")

    # again we have bytes where we need strings. :-(
    # could have used email.message_from_bytes ???

    msg_str = msg.decode('UTF-8')
    # print(f"Message is\n----------\n{msg_str}\n----------\n")
    msgobj = email.message_from_string(msg_str)  # was msg

    # msgobj = email.message_from_bytes(msg)

    to_info = split_addrs(msgobj['to'])

    email_name = to_info[0][0]
    email_addr = to_info[0][1]
    print(f" name: '{email_name}', addr: '{email_addr}'")
    # old_item = addrs_with_names[]
    if len(email_name) == 0:
        print("   skipping empty name")
        skipped += 1
        continue
    if email_name in names_and_addresses:
        addresses = names_and_addresses[email_name]
    else:
        print("  new name!")
        addresses = set()
        names_and_addresses[email_name] = addresses
    print(f"  Prev set of addresses is {addresses}")
    addresses.add(email_addr)
    if len(addresses) == 4:
        print("  **** MAX NAMES HIT FOR {email_name}")

    # if iter == 10:
    #     print("exiting early")
    #     print(f"addrs_with_names: {addrs_with_names}")
    #     exit(1)


print(f"Found {len(names_and_addresses)} distinct names:")
print(f"{names_and_addresses}")

# Now create a list of dictionaries for the CSV DictWriter.
# Each entry in the list is a dictionary like this:
#  {"Name": "Rob Cranfill",
#   "E-mail 1 - Value": "robcranfill@comcast.net",
#   "E-mail 2 - Value": "robcranfill@gmail.com"}
#
csv_data = []
for na in names_and_addresses.keys():
    d = dict()
    d["Name"] = na
    i = 1
    # todo: check for email addresses that only differ in upper/lower case?
    for a in names_and_addresses[na]:
        key_name = f"E-mail {i} - Value"
        d[key_name] = a
        print(f" >>> {key_name} = {a}")
        i += 1
    csv_data.append(d)

print(f"\n\nCSV dict: {csv_data}")

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
    for d in csv_data:
        writer.writerow(d)

import imaplib
import email
import sys
# extract headers from online email!
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

# this is the final list; don't we want to have only distinct entries???
addr = []

# more-useful result: dictionary of
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
    # print(f"to_info = {to_info}")
    # exit(2)

    email_name = to_info[0][0]
    email_addr = to_info[0][1]
    print(f"to_info: name: '{email_name}', addr: '{email_addr}'")
    # old_item = addrs_with_names[]
    if len(email_name) == 0:
        print("skipping empty name")
        skipped += 1
        continue
    if email_name in names_and_addresses:
        addresses = names_and_addresses[email_name]
    else:
        print("  new name!")
        addresses = set()
        names_and_addresses[email_name] = addresses
    print(f"Prev set of addresses is {addresses}")
    addresses.add(email_addr)
    if len(addresses) == 4:
        print("MAX NAMES HIT FOR {email_name}")

    # if iter == 10:
    #     print("exiting early")
    #     print(f"addrs_with_names: {addrs_with_names}")
    #     exit(1)

    addr.extend(to_info)    # obsolete
    addr.extend(split_addrs(msgobj['from']))
    addr.extend(split_addrs(msgobj['cc']))

print(f"Found {len(addr)} addresses:")
print(f"{addr}")

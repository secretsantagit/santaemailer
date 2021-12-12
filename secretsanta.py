import argparse
import json
import random
import time
import win32com.client

if __name__ == '__main__':
    # parsing arguments
    parser = argparse.ArgumentParser(description="This is a Secret Santa emailer v1.0. It builds randomly the pair of people (2+) and sends the emails via your MS Outlook.")
    parser.add_argument("-c", "--configfile", help="Path to a file with settings.", required=True)
    parser.add_argument("-t", "--test", help="Just test, but not send the email.", action='store_true')
    parser.add_argument("-t2", "--test2", help="Just show a list of people pairs.", action='store_true')
    args = parser.parse_args()

    # loading config
    with open(args.configfile) as json_file:
        cfg = json.load(json_file)

    # reading people form the config
    contacts_from = []
    for man in cfg["people"]:
        contacts_from.append(man)
    if len(contacts_from) < 2:
        print ("We need more than 1 man")
        exit(0)

    # preparation of a random pairs of people for the mailing

    while True:
        tmp = contacts_from.copy()
        contacts_to = []
        f = 0
        while len(tmp) > 0:
            random.seed(time.time_ns())
            t = random.randint(0, len(tmp) - 1)
            if contacts_from[f] == tmp[t]:
                if len(tmp) == 1:
                    break  # we are in the dead end
                continue  # skip a case if Santa sends a present himself
            contacts_to.append(tmp.pop(t))
            f += 1
        if len(tmp) == 0:
            break

    # message template preparation
    msg = ""
    for row in cfg["message"]:
        msg += f"{row}\n"

    # preparation and sending of an email
    i = 0
    print ("Let's start...\n")
    while i < len(contacts_from):
        # message preparation
        m = msg[:]
        m = m.replace("%DEARFROM%", contacts_from[i]["name"])
        m = m.replace("%DEARTO%", contacts_to[i]["name"])

        # email sending
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = contacts_from[i]["email"]
        mail.Subject = cfg["subject"]
        mail.Body = m
        if args.test:
            print (f"To: {mail.To}\nSubject: {mail.Subject}\n")
            print (mail.Body)
            print ("-------------------------------- test message, has not been sent")
        elif args.test2:
            print(f"from {contacts_from[i]['name']} fo {contacts_to[i]['name']}")
        else:
            mail.Send()
            print(f"{i+1} email from {len(contacts_from)} has been sent")
        i += 1

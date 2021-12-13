import argparse
import yaml
import random
import time
import win32com.client

if __name__ == '__main__':
    # parsing arguments
    parser = argparse.ArgumentParser(description="This is a Secret Santa emailer v1.1. It builds randomly the pair of people (2+) and sends the emails via your MS Outlook.")
    parser.add_argument("-c", "--configfile", help="Path to a file with settings.", required=True)
    parser.add_argument("-t", "--test_num", type=int, help="t=1: Just show a list of people pairs. t=2: Show emails, but not send them.")
    args = parser.parse_args()

    # loading config
    with open(args.configfile) as file:
        cfg = yaml.load(file, Loader=yaml.FullLoader)

    # reading people form the config
    contacts = cfg["people"].copy()
    random.shuffle(contacts)

    if len(contacts) < 2:
        print ("We need more than 1 man")
        exit(0)

    msg = cfg["message"]

    print ("Let's start...\n")
    i = 0
    while i < len(contacts):
        # message preparation
        m = msg[:]
        m = m.replace("%DEARFROM%", contacts[i]["name"])
        m = m.replace("%DEARTO%", contacts[i-1]["name"])

        # email sending
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = contacts[i]["email"]
        mail.Subject = cfg["subject"]
        mail.Body = m
        if args.test_num == 1:
            print(f"from {contacts[i]['name']} fo {contacts[i-1]['name']}")
        elif args.test_num == 2:
            print(f"To: {mail.To}\nSubject: {mail.Subject}\n")
            print(mail.Body)
            print("-------------------------------- test message, has not been sent")
        else:
            mail.Send()
            print(f"{i+1} email from {len(contacts)} has been sent")
        i += 1

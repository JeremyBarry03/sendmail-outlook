
import argparse
import sys
import os
import tempfile
from win32com.client import Dispatch

def create_mail(outlook, to, subject, body, account=None):
    mail = outlook.CreateItem(0)
    if account is not None:
        mail.SendUsingAccount = account
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    return mail

def create_parser():
    parser = argparse.ArgumentParser(description="Sends an email using Outlook")
    parser.add_argument("-t", "--to", default="", help="Recipient of the email")
    parser.add_argument(
        "-s", "--subject", default="Message subject", help="Subject of the email"
    )
    parser.add_argument(
        "-b", "--body", default="Message body", help="Body of the email", nargs="*"
    )
    parser.add_argument(
        "-hb", "--htmlbody", default=None, help="HTML Body of the email"
    )
    parser.add_argument(
        "-a",
        "--attachment",
        nargs="?",
        const=None,
        type=str,
        help="Path to attachment file (optional)",
    )

    parser.add_argument("-ds", "--dont-send", action="store_true", help="Don't send email, only show")    

    return parser

def main():
    parser = create_parser()
    args = parser.parse_args()

    dont_send = args.dont_send

    # If no to is provided, don't send the email just display it.
    if len(args.to) == 0:
        dont_send = True
        
    outlook = Dispatch("outlook.application")
    if not args.body:
        body = sys.stdin.read()
    else:
        body = args.body if isinstance(args.body, str) else " ".join(args.body)
    mail = create_mail(outlook, args.to, args.subject, body)

    if args.htmlbody:
        mail.HTMLBody = args.htmlbody

    if args.attachment:
        attachment_path = args.attachment
        mail.Attachments.Add(attachment_path)

    if dont_send:
        temp_dir = tempfile.gettempdir()    
        msg_file_path = os.path.join(temp_dir, "email.msg")
        mail.SaveAs(msg_file_path, 3)
        os.startfile(msg_file_path)
    else:
        mail.Send()


if __name__ == "__main__":
    main()

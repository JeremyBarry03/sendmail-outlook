# sendmail-outlook

Send personalized emails through classic Outlook on Windows.

## Desktop app

Requires Windows, Python 3.8+ with Tkinter, and **classic Outlook** with your
sending account configured. This app uses classic Outlook's existing login.
No Microsoft app registration, client ID, client secret, or API keys are needed.
New Outlook does not support the COM automation used here.

Install this checkout and launch:

```powershell
python -m pip install -e .
python outlook_gui.py
```

After installation, `sendmail-gui` also launches the window. You can double-click
`launch_gui.pyw` if Windows associates `.pyw` with your Python installation.
Close any older GUI window before launching the updated version.

1. Open **Outlook (classic)** and sign into your sending account.
2. Click **Connect Outlook**, then choose the **From** account.
3. Enter one recipient in **To**, or click **Add CSV...** and select the email column.
4. Write a subject and message. Insert CSV columns using exact headers in braces,
   such as `{name}`, `{company}`, or `{First Name}`. Names are case-sensitive.
5. Click **Preview emails** and inspect the personalized messages.
6. Click **Send emails** and confirm the account and recipient count.

For example, use `example-recipients.csv` with this message:

```text
Hi {name},

I wanted to reach out about {company}.
```

Joe receives "Hi Joe" and Jeremy receives "Hi Jeremy". Each CSV row becomes a
separate plain-text email. The CSV replaces the manual To field while loaded.
Export CSV as UTF-8 (Excel's "CSV UTF-8" option), with unique column headers.
Use `{{` and `}}` for literal braces. CSV import uses a file picker.

## Sending and stopping

All rows are validated before sending. Malformed addresses, unknown placeholders,
and duplicate recipient addresses stop the batch before any email is submitted.
The sender and message fields are locked while sending. An Outlook error stops
the batch without an automatic retry.

**Stop after current email** stops the remaining rows. COM calls run on the GUI
thread, so Outlook may temporarily pause the window during a send or a security
prompt; the Stop button takes effect once control returns. It cannot recall an
email or remove messages already queued in Outlook's Outbox. Closing during a
batch requests a stop; close again after it stops.

**Save results...** exports attempted recipients and their status for this session.
"Submitted to Outlook" means Outlook accepted the request, not that delivery was
confirmed. Check Outbox/Sent Items and any delivery failure reports. Unattempted
rows are not in the report. Identical sender/recipient/subject/body combinations
cannot be attempted twice in the same session. This protection does not persist
after restarting: check results and Outlook before resending. Replace the sample
CSV's example addresses before a real send.

## Authentication and account selection

The app uses `pywin32` / Windows COM to control the locally installed classic
Outlook application. Outlook handles account login and MFA. The app does not
collect or store passwords or tokens, and no environment variables are required.
**Connect Outlook** lists sending accounts in Outlook's active profile; the GUI
sets `SendUsingAccount` on every message to the selected account. The original
CLI leaves account selection to Outlook.

Being signed into new Outlook does not configure classic Outlook. If no accounts
are found, open classic Outlook and add the account there, then reconnect.
School/work policies still apply: the account must be permitted to sign into
classic Outlook and send through automation. Shared mailboxes and aliases are not
automatically separate selectable accounts.

The previous Graph connection and registration screen have been removed. Any
previously saved `%LOCALAPPDATA%\OutlookMailer\settings.json` is no longer read.

## Development checks

```powershell
python -m pip install -e ".[test]"
python -m pytest -q
python -m compileall -q send_outlookemail.py mail_merge.py outlook_gui.py
python -m build
```

Tests require a Windows desktop session and mock Outlook; they do not send real
emails. A live send through your account has to be tested separately.

## Classic Outlook CLI usage

The original CLI requires Windows and classic Outlook configured with an account.
```
usage: sendmail [-h] -t TO [-s SUBJECT] [-b [BODY ...]] [-hb HTMLBODY] [-a [ATTACHMENT]]

Sends an email using Outlook

options:
  -h, --help            show this help message and exit
  -t TO, --to TO        Recipient of the email
  -s SUBJECT, --subject SUBJECT
                        Subject of the email
  -b [BODY ...], --body [BODY ...]
                        Body of the email
  -hb HTMLBODY, --htmlbody HTMLBODY
                        HTML Body of the email
  -a [ATTACHMENT], --attachment [ATTACHMENT]
                        Path to attachment file (optional)
```

# Install
pip install sendmail-outlook

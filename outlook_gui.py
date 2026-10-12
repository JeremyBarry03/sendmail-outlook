"""Desktop mail merge using the signed-in classic Outlook profile."""

import csv
import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from win32com.client import Dispatch
from send_outlookemail import create_mail
from mail_merge import prepare_messages, read_csv


class Mailer:
    def __init__(self, root):
        self.root = root
        self.outlook = None
        self.accounts = []
        self.rows = None
        self.busy = False
        self.cancelled = False
        self.attempted = set()
        self.results = []
        root.title("Outlook mailer")
        root.geometry("850x800")
        root.minsize(620, 620)
        root.protocol("WM_DELETE_WINDOW", self.close)
        self.form = ttk.Frame(root, padding=16)
        self.form.pack(fill="both", expand=True)
        self.form.columnconfigure(1, weight=1)
        ttk.Label(self.form, text="Send personalized emails", font=("Segoe UI", 16)).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))
        ttk.Label(self.form, text="From").grid(row=1, column=0, sticky="w")
        self.sender = ttk.Combobox(self.form, state="readonly")
        self.sender.grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(self.form, text="Connect Outlook", command=self.connect).grid(row=1, column=2)
        ttk.Label(self.form, text="To (without CSV)").grid(row=2, column=0, sticky="w", pady=10)
        self.to = ttk.Entry(self.form)
        self.to.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(8, 0))
        csv_bar = ttk.Frame(self.form)
        csv_bar.grid(row=3, column=0, columnspan=3, sticky="ew")
        ttk.Button(csv_bar, text="Add CSV…", command=self.load_csv).pack(side="left")
        ttk.Button(csv_bar, text="Remove CSV", command=self.clear_csv).pack(side="left", padx=8)
        self.csv_label = ttk.Label(csv_bar, text="No CSV selected")
        self.csv_label.pack(side="left")
        ttk.Label(self.form, text="Email column").grid(row=4, column=0, sticky="w", pady=10)
        self.email_column = ttk.Combobox(self.form, state="readonly")
        self.email_column.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(8, 0))
        self.columns_label = ttk.Label(self.form, text="Use {name} to insert a CSV column. Names are case-sensitive.", wraplength=580)
        self.columns_label.grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 12))
        ttk.Label(self.form, text="Subject").grid(row=6, column=0, sticky="w")
        self.subject = ttk.Entry(self.form)
        self.subject.grid(row=6, column=1, columnspan=2, sticky="ew", padx=(8, 0))
        ttk.Label(self.form, text="Message").grid(row=7, column=0, sticky="w", pady=(12, 4))
        self.body = ScrolledText(self.form, height=6, width=45, wrap="word", font=("Segoe UI", 11))
        self.body.grid(row=8, column=0, columnspan=3, sticky="nsew")
        self.form.rowconfigure(8, weight=1)
        buttons = ttk.Frame(self.form)
        buttons.grid(row=9, column=0, columnspan=3, sticky="w", pady=12)
        ttk.Button(buttons, text="Preview emails", command=self.preview).pack(side="left")
        ttk.Button(buttons, text="Send emails", command=self.send).pack(side="left", padx=8)
        ttk.Button(buttons, text="Save results…", command=self.save_results).pack(side="left")
        self.status = tk.StringVar(value="Open classic Outlook and sign in, then click Connect Outlook.")
        ttk.Label(root, textvariable=self.status, wraplength=580, padding=(16, 4)).pack(anchor="w")
        self.stop = ttk.Button(root, text="Stop after current email", command=self.cancel, state="disabled")
        self.stop.pack(anchor="w", padx=16, pady=(0, 12))
        root.update_idletasks()
        root.minsize(max(620, root.winfo_reqwidth()), max(620, root.winfo_reqheight()))

    def connect(self):
        if self.busy:
            return
        try:
            outlook = Dispatch("Outlook.Application")
            accounts = [account for account in outlook.Session.Accounts if account.SmtpAddress]
            if not accounts:
                raise ValueError("No sending accounts found. Add your account in classic Outlook first.")
            labels = [f"{account.DisplayName} <{account.SmtpAddress}>" for account in accounts]
        except Exception as error:
            self.outlook, self.accounts = None, []
            self.sender.set("")
            self.sender.configure(values=[])
            self.status.set("Not connected. Open classic Outlook and sign in, then try again.")
            messagebox.showerror("Could not connect", f"This app requires classic Outlook, not new Outlook.\n\n{error}", parent=self.root)
            return
        self.outlook, self.accounts = outlook, accounts
        self.sender.configure(values=labels)
        self.sender.current(0)
        self.status.set("Connected to classic Outlook. Check the From account before sending.")

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")], parent=self.root)
        if not path:
            return
        try:
            columns, rows = read_csv(path)
        except (OSError, UnicodeError, ValueError, csv.Error) as error:
            messagebox.showerror("Invalid CSV", str(error), parent=self.root)
            return
        self.rows = rows
        self.csv_label.configure(text=f"{Path(path).name[:35]} · {len(rows)} recipients")
        self.email_column.configure(values=columns)
        selected = next((name for name in columns if name.lower() in ("email", "to", "email_address")), columns[0])
        self.email_column.set(selected)
        self.columns_label.configure(text="Available placeholders: " + ", ".join("{" + name + "}" for name in columns))
        self.to.configure(state="disabled")

    def clear_csv(self):
        self.rows = None
        self.csv_label.configure(text="No CSV selected")
        self.email_column.set("")
        self.email_column.configure(values=[])
        self.columns_label.configure(text="Use {name} to insert a CSV column. Names are case-sensitive.")
        self.to.configure(state="normal")

    def messages(self):
        rows = self.rows if self.rows is not None else [{"email": self.to.get()}]
        column = self.email_column.get() if self.rows is not None else "email"
        return prepare_messages(rows, column, self.subject.get(), self.body.get("1.0", "end-1c"))

    def preview(self):
        try:
            messages = self.messages()
        except ValueError as error:
            messagebox.showerror("Check your message", str(error), parent=self.root)
            return
        window = tk.Toplevel(self.root)
        window.title("Email preview — nothing sent")
        window.geometry("640x500")
        selector = ttk.Combobox(window, state="readonly", values=[f"{i + 1}. {mail.to}" for i, mail in enumerate(messages)])
        selector.pack(fill="x", padx=12, pady=12)
        content = ScrolledText(window, wrap="word")
        content.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        def show(_event=None):
            mail = messages[selector.current()]
            content.configure(state="normal")
            content.delete("1.0", "end")
            content.insert("1.0", f"From: {self.sender.get() or '(choose an Outlook account)'}\nTo: {mail.to}\nSubject: {mail.subject}\n\n{mail.body}")
            content.configure(state="disabled")

        selector.bind("<<ComboboxSelected>>", show)
        selector.current(0)
        show()

    def set_busy(self, busy):
        self.busy = busy

        def toggle(widget):
            for child in widget.winfo_children():
                if isinstance(child, (ttk.Button, ttk.Entry, ttk.Combobox, tk.Text)):
                    state = "disabled" if busy else "readonly" if isinstance(child, ttk.Combobox) else "normal"
                    child.configure(state=state)
                toggle(child)

        toggle(self.form)
        if self.rows is not None:
            self.to.configure(state="disabled")
        self.stop.configure(state="normal" if busy else "disabled")

    def send(self):
        if self.busy:
            return
        try:
            messages = self.messages()
            if self.outlook is None or self.sender.current() < 0:
                raise ValueError("Connect Outlook and select a sending account first.")
            account = self.accounts[self.sender.current()]
            sender = account.SmtpAddress
            if any((sender.casefold(), mail) in self.attempted for mail in messages):
                raise ValueError("This session already attempted one or more of these emails. Check the results and Sent Items, then remove those rows before continuing.")
        except Exception as error:
            messagebox.showerror("Cannot send", str(error), parent=self.root)
            return
        if not messagebox.askyesno("Send emails?", f"Send {len(messages)} separate emails from {sender}?\n\nEach recipient receives only their personalized email.", parent=self.root):
            return
        self.pending = iter(messages)
        self.batch_account = account
        self.batch_sender = sender
        self.total = len(messages)
        self.sent = 0
        self.cancelled = False
        self.set_busy(True)
        self.root.after(10, self.send_next)

    def send_next(self):
        if self.cancelled:
            self.finish(f"Stopped. {self.sent}/{self.total} emails submitted to Outlook.")
            return
        message = next(self.pending, None)
        if message is None:
            self.finish(f"Done. {self.sent}/{self.total} emails submitted to Outlook. Check Outbox/Sent Items; delivery is not yet confirmed.")
            return
        self.status.set(f"Sending email {self.sent + 1}/{self.total} through classic Outlook…")
        self.attempted.add((self.batch_sender.casefold(), message))
        try:
            mail = create_mail(self.outlook, message.to, message.subject, message.body, self.batch_account)
            mail.Send()
        except Exception as error:
            self.results.append((self.batch_sender, message.to, "Not confirmed; check Outlook before retrying", str(error)))
            self.finish(f"Stopped after an error. {self.sent}/{self.total} submitted. Save results for details.")
            messagebox.showerror("Sending stopped", f"Could not confirm sending to {message.to}.\n\n{error}\n\nCheck Sent Items before retrying; no automatic retry was made.", parent=self.root)
            return
        self.sent += 1
        self.results.append((self.batch_sender, message.to, "Submitted to Outlook (delivery not confirmed)", ""))
        self.root.after(100, self.send_next)

    def finish(self, text):
        self.set_busy(False)
        self.status.set(text)

    def cancel(self):
        self.cancelled = True

    def close(self):
        if self.busy:
            self.cancel()
            self.status.set("Stopping after the current email. Close the window once stopped.")
        else:
            self.root.destroy()

    def save_results(self):
        if not self.results:
            messagebox.showinfo("Results", "No emails have been attempted in this session.", parent=self.root)
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="send-results.csv", filetypes=[("CSV files", "*.csv")], parent=self.root)
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as target:
                writer = csv.writer(target)
                writer.writerow(["sender", "recipient", "status", "error"])
                for row in self.results:
                    writer.writerow(["'" + value if value.startswith(("=", "+", "-", "@", "\t", "\r", "\n")) else value for value in row])
        except OSError as error:
            messagebox.showerror("Could not save results", str(error), parent=self.root)


def main():
    if os.name == "nt":
        # Keep the full window visible and text crisp with Windows display scaling.
        import ctypes
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass
    root = tk.Tk()
    Mailer(root)
    root.mainloop()


if __name__ == "__main__":
    main()

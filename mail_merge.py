"""CSV validation and literal column substitution for the desktop mailer."""

import csv
import re
from dataclasses import dataclass
from string import Formatter


@dataclass(frozen=True)
class Message:
    to: str
    subject: str
    body: str


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source, strict=True)
        columns = [name.strip() for name in (reader.fieldnames or [])]
        if not columns or any(not name or any(c in name for c in "{}") for name in columns):
            raise ValueError("CSV needs a header row with nonempty column names without braces.")
        if len(set(columns)) != len(columns):
            raise ValueError("CSV column names must be unique.")
        reader.fieldnames = columns
        rows = []
        for row in reader:
            if None in row or any(value is None for value in row.values()):
                raise ValueError(f"CSV line {reader.line_num} has the wrong number of columns.")
            if any(value.strip() for value in row.values()):
                rows.append({key: value.strip() for key, value in row.items()})
        if not rows:
            raise ValueError("CSV has no recipients.")
        return columns, rows


def render(template, row):
    """Allow exact column names only, never attribute access or expressions."""
    parts = []
    for literal, field, spec, conversion in Formatter().parse(template):
        parts.append(literal)
        if field is not None:
            if field not in row:
                raise ValueError(f"Unknown placeholder {{{field}}}. Use a CSV column name exactly.")
            if spec or conversion:
                raise ValueError("Use plain placeholders such as {name}, without formatting codes.")
            parts.append(row[field])
    return "".join(parts)


def prepare_messages(rows, email_column, subject, body):
    if not subject.strip() or not body.strip():
        raise ValueError("Enter both a subject and a message.")
    messages = []
    seen = set()
    for number, row in enumerate(rows, 1):
        address = row.get(email_column, "").strip()
        if not re.fullmatch(r"[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+", address):
            raise ValueError(f"Recipient row {number}: enter one valid email address in '{email_column}'.")
        if address.casefold() in seen:
            raise ValueError(f"Recipient row {number}: duplicate email address {address}.")
        seen.add(address.casefold())
        try:
            message = Message(address, render(subject, row), render(body, row))
        except ValueError as error:
            raise ValueError(f"Recipient row {number}: {error}") from error
        if not message.subject.strip() or not message.body.strip():
            raise ValueError(f"Recipient row {number}: personalized subject or message is empty.")
        if "\n" in message.subject or "\r" in message.subject:
            raise ValueError(f"Recipient row {number}: subject must be a single line.")
        messages.append(message)
    if not messages:
        raise ValueError("Add at least one recipient.")
    return messages

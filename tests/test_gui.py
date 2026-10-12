from unittest.mock import Mock, patch
import tkinter as tk

import pytest

from outlook_gui import Mailer
from send_outlookemail import create_mail


@pytest.fixture(scope="module")
def root():
    root = tk.Tk()
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture
def app(root):
    instance = Mailer(root)
    yield instance
    for child in root.winfo_children():
        child.destroy()


def configured(app):
    app.outlook = Mock()
    account = Mock(SmtpAddress="sender@example.com")
    app.accounts = [account]
    app.sender.configure(values=["sender@example.com"])
    app.sender.current(0)
    app.rows = [{"email": "joe@example.com", "name": "Joe"}, {"email": "jeremy@example.com", "name": "Jeremy"}]
    app.email_column.set("email")
    app.subject.insert(0, "Hi {name}")
    app.body.insert("1.0", "Hello {name}")
    return app.outlook


def test_shared_mail_uses_explicit_sender():
    outlook, account = Mock(), Mock()
    mail = create_mail(outlook, "joe@example.com", "Hi", "Hello", account)
    assert mail.SendUsingAccount is account
    assert (mail.To, mail.Subject, mail.Body) == ("joe@example.com", "Hi", "Hello")
    mail.Send.assert_not_called()


def test_send_personalized_and_block_repeat(app):
    outlook = configured(app)
    first, second = Mock(), Mock()
    outlook.CreateItem.side_effect = [first, second]
    with patch("outlook_gui.messagebox.askyesno", return_value=True), patch.object(app.root, "after"):
        app.send()
        app.send_next()
        app.send_next()
        app.send_next()
    assert first.Body == "Hello Joe"
    assert second.Body == "Hello Jeremy"
    assert first.SendUsingAccount is app.accounts[0]
    assert second.SendUsingAccount is app.accounts[0]
    first.Send.assert_called_once()
    second.Send.assert_called_once()
    assert not app.busy
    with patch("outlook_gui.messagebox.showerror") as error:
        app.send()
        error.assert_called_once()
    assert outlook.CreateItem.call_count == 2


def test_failure_stops_batch_without_retry(app):
    outlook = configured(app)
    outlook.CreateItem.return_value.Send.side_effect = RuntimeError("Connection lost")
    with patch("outlook_gui.messagebox.askyesno", return_value=True), patch("outlook_gui.messagebox.showerror"), patch.object(app.root, "after") as schedule:
        app.send()
        schedule.reset_mock()
        app.send_next()
        schedule.assert_not_called()
    assert not app.busy
    assert len(app.attempted) == 1
    assert "Not confirmed" in app.results[0][2]
    outlook.CreateItem.return_value.Send.assert_called_once()


def test_cancel_and_preview_never_send(app):
    outlook = configured(app)
    app.preview()
    outlook.CreateItem.assert_not_called()
    with patch("outlook_gui.messagebox.askyesno", return_value=True), patch.object(app.root, "after"):
        app.send()
        app.cancel()
        app.send_next()
    outlook.CreateItem.assert_not_called()
    assert not app.busy


def test_bad_later_row_prevents_entire_batch(app):
    outlook = configured(app)
    app.rows[1]["email"] = "bad"
    with patch("outlook_gui.messagebox.showerror"), patch("outlook_gui.messagebox.askyesno") as confirm:
        app.send()
    confirm.assert_not_called()
    outlook.CreateItem.assert_not_called()


def test_connect_lists_accounts_without_sending(app):
    first = Mock(DisplayName="School", SmtpAddress="school@example.com")
    second = Mock(DisplayName="Personal", SmtpAddress="personal@example.com")
    with patch("outlook_gui.Dispatch") as dispatch:
        dispatch.return_value.Session.Accounts = [first, second]
        app.connect()
        assert app.accounts == [first, second]
        assert "school@example.com" in app.sender.get()
        dispatch.return_value.CreateItem.assert_not_called()


def test_failed_reconnect_clears_stale_accounts(app):
    configured(app)
    with patch("outlook_gui.Dispatch") as dispatch, patch("outlook_gui.messagebox.showerror"):
        dispatch.return_value.Session.Accounts = []
        app.connect()
    assert app.outlook is None
    assert app.accounts == []
    assert app.sender.get() == ""


def test_selected_sender_is_used_and_locked_during_batch(app):
    outlook = configured(app)
    second = Mock(SmtpAddress="second@example.com")
    app.accounts.append(second)
    app.sender.configure(values=["sender@example.com", "second@example.com"])
    app.sender.current(1)
    with patch("outlook_gui.messagebox.askyesno", return_value=True), patch.object(app.root, "after"), patch("outlook_gui.Dispatch") as dispatch:
        app.send()
        assert str(app.sender.cget("state")) == "disabled"
        app.connect()
        dispatch.assert_not_called()
        app.send_next()
        assert outlook.CreateItem.return_value.SendUsingAccount is second
        app.cancel()
        app.send_next()
    assert app.sent == 1
    assert not app.busy

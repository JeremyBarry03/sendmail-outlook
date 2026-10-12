import pytest

from mail_merge import prepare_messages, read_csv, render


def test_excel_csv_and_personalization(tmp_path):
    path = tmp_path / "people.csv"
    path.write_text('email,name,company\njoe@example.com,Joe,"Acme, Inc"\njeremy@example.com,Jeremy,Example\n', encoding="utf-8-sig")
    columns, rows = read_csv(path)
    messages = prepare_messages(rows, "email", "Hello {name}", "Hi {name},\nHow is {company}?")
    assert columns == ["email", "name", "company"]
    assert messages[0].body == "Hi Joe,\nHow is Acme, Inc?"
    assert messages[1].subject == "Hello Jeremy"


@pytest.mark.parametrize("content", ["", "email,name\n", "email,email\na,b\n", "email,name\na\n", "email,name\na,b,c\n", "email,\na,b\n"])
def test_reject_malformed_csv(tmp_path, content):
    path = tmp_path / "bad.csv"
    path.write_text(content, encoding="utf-8")
    with pytest.raises(ValueError):
        read_csv(path)


@pytest.mark.parametrize("template", ["Hi {missing}", "Hi {name", "Hi {name.upper}", "Hi {name!r}"])
def test_reject_bad_templates(template):
    with pytest.raises(ValueError):
        render(template, {"name": "Joe"})


def test_literal_braces_and_values_not_reinterpreted():
    assert render("{{hello}} {first name}", {"first name": "{Joe}"}) == "{hello} {Joe}"


@pytest.mark.parametrize("addresses", [["bad"], ["a@example.com;b@example.com"], ["a@example.com", "A@example.com"]])
def test_reject_invalid_or_duplicate_recipients(addresses):
    with pytest.raises(ValueError):
        prepare_messages([{"email": address} for address in addresses], "email", "Hi", "Hello")


def test_validate_all_rows_before_returning():
    with pytest.raises(ValueError, match="row 2"):
        prepare_messages([{"email": "joe@example.com", "name": "Joe"}, {"email": "bad", "name": "Jeremy"}], "email", "Hi {name}", "Hello")

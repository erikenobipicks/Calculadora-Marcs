"""Tests dels helpers purs del mailing (parsing, validació, render).
La part de DB (esquema, upsert, dedupe) es valida a part contra SQLite."""
import app


def test_valid_email():
    assert app._mailing_valid_email("A@B.com") == "a@b.com"
    assert app._mailing_valid_email("  Foo@Bar.CAT ") == "foo@bar.cat"
    assert app._mailing_valid_email("nope") == ""
    assert app._mailing_valid_email("x@y") == ""        # sense punt al domini
    assert app._mailing_valid_email("") == ""
    assert app._mailing_valid_email(None) == ""


def test_parse_lines_formats():
    txt = (
        "anna@x.com\n"
        "Joan Pla, joan@x.com\n"
        "Maria <maria@x.com>\n"
        "Pere; pere@x.com\n"
        "linia-sense-mail\n"
        "\n"
    )
    out = app._mailing_parse_lines(txt)
    assert ("", "anna@x.com") in out
    assert ("Joan Pla", "joan@x.com") in out
    assert ("Maria", "maria@x.com") in out
    assert ("Pere", "pere@x.com") in out
    # la línia sense mail i la buida no generen contacte
    assert all("@" in e for _, e in out)
    assert len(out) == 4


def test_text_to_html_escapes_and_paragraphs():
    html = app._mailing_text_to_html("Hola {nom}\n\nSegon <b>raw</b> paràgraf")
    assert "&lt;b&gt;" in html          # l'HTML cru queda escapat
    assert html.count("<p ") == 2        # dos paràgrafs
    assert "{nom}" in html               # el marcador es manté per substituir després


def test_render_html_has_unsubscribe_and_name(monkeypatch):
    monkeypatch.setattr(app, "get_config_value", lambda clau, default=None: default)
    contact = {"nom": "Anna", "token": "TOK123"}
    body = app._mailing_text_to_html("Hola {nom}")
    html = app._mailing_render_html(body, contact)
    assert "Anna" in html                # {nom} substituït
    assert "/baixa/TOK123" in html       # enllaç de baixa tokenitzat

"""Tests de la secció admin de Consumibles (registre + comanda per email).

Cobreix la normalització, el format de quantitat/escapat, la generació del cos
del correu (amb escapat), l'agrupació de pendents per equip amb proveïdor, i que
la pàgina és només accessible per admin.
"""
import pytest

import app


def test_normalize_consumible_basic():
    c = app._normalize_consumible({'nom': '  Cartutx cian ', 'equip': 'noritsu_green_iv',
                                   'quantitat': '2', 'pendent': 1})
    assert c['nom'] == 'Cartutx cian'
    assert c['equip'] == 'noritsu_green_iv'
    assert c['quantitat'] == 2.0
    assert c['pendent'] is True
    assert c['actiu'] is True


def test_normalize_consumible_unknown_equip_falls_back():
    assert app._normalize_consumible({'nom': 'X', 'equip': 'inexistent'})['equip'] == 'altres'


def test_normalize_consumible_bad_qty():
    assert app._normalize_consumible({'nom': 'X', 'quantitat': 'abc'})['quantitat'] == 0


def test_consum_qty_formatting():
    assert app._consum_qty(2) == '2'
    assert app._consum_qty(2.0) == '2'
    assert app._consum_qty(2.5) == '2.50'
    assert app._consum_qty('bad') == '0'


def test_consum_esc():
    assert app._consum_esc('<b>&') == '&lt;b&gt;&amp;'
    assert app._consum_esc(None) == ''


def test_email_html_lists_items_and_escapes():
    grup = {
        'equip_label': 'Noritsu Green IV',
        'proveidor_nom': 'Prov <X>',
        'items': [
            {'nom': 'Cartutx cian', 'referencia': 'ABC-1', 'quantitat': 2},
            {'nom': 'Paper <lustre>', 'referencia': '', 'quantitat': 1},
        ],
    }
    html = app._consumibles_email_html(grup)
    assert 'Noritsu Green IV' in html
    assert 'Cartutx cian' in html
    assert 'ABC-1' in html
    assert '&lt;lustre&gt;' in html          # nom escapat
    assert '<lustre>' not in html            # cap injecció crua
    assert 'Prov &lt;X&gt;' in html          # proveïdor escapat
    assert '—' in html                       # referència buida → guió


def test_comanda_grups_groups_pending(monkeypatch):
    items = [
        {'equip': 'noritsu_green_iv', 'nom': 'Cian', 'referencia': '', 'quantitat': 1, 'pendent': True,  'notes': '', 'actiu': True, 'ordre': 1},
        {'equip': 'noritsu_green_iv', 'nom': 'Negre', 'referencia': '', 'quantitat': 1, 'pendent': False, 'notes': '', 'actiu': True, 'ordre': 2},
        {'equip': 'canon_pro_4000',   'nom': 'Tinta', 'referencia': '', 'quantitat': 3, 'pendent': True,  'notes': '', 'actiu': True, 'ordre': 3},
    ]
    monkeypatch.setattr(app, 'get_consumibles_list', lambda: items)
    monkeypatch.setattr(app, 'get_config_value',
                        lambda clau, default=None: ('x@prov.com' if clau.endswith('_email')
                                                    else 'ProvX' if clau.endswith('_nom') else default))
    grups = app._consumibles_comanda_grups()
    # Només equips amb pendents, en l'ordre d'EQUIPS_CONSUMIBLES (canon abans que noritsu).
    assert [g['equip_key'] for g in grups] == ['canon_pro_4000', 'noritsu_green_iv']
    noritsu = next(g for g in grups if g['equip_key'] == 'noritsu_green_iv')
    assert len(noritsu['items']) == 1  # només el pendent, no el 'Negre'
    assert noritsu['proveidor_email'] == 'x@prov.com'
    assert noritsu['proveidor_nom'] == 'ProvX'


def test_comanda_grups_empty_when_none_pending(monkeypatch):
    monkeypatch.setattr(app, 'get_consumibles_list',
                        lambda: [{'equip': 'canon_pro_4000', 'nom': 'X', 'quantitat': 1,
                                  'pendent': False, 'actiu': True, 'referencia': '', 'notes': '', 'ordre': 1}])
    assert app._consumibles_comanda_grups() == []


def test_admin_consumibles_requires_admin():
    app.app.config['TESTING'] = True
    client = app.app.test_client()
    r = client.get('/admin/consumibles')
    assert r.status_code != 200  # gated: sense sessió admin no s'hi accedeix

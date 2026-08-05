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


# ── Preu i històric de preus ───────────────────────────────────────────────
def test_normalize_includes_preu_id_historial():
    c = app._normalize_consumible({'nom': 'Cian', 'preu': '18.5', 'id': 'abc',
                                   'historial': [{'preu': '15', 'data': '2026-01-01'}]})
    assert c['id'] == 'abc'
    assert c['preu'] == 18.5
    assert c['historial'] == [{'preu': 15.0, 'data': '2026-01-01'}]


def test_normalize_bad_preu_and_historial():
    c = app._normalize_consumible({'nom': 'X', 'preu': 'abc', 'historial': 'nope'})
    assert c['preu'] == 0.0
    assert c['historial'] == []


def test_apply_historial_new_item_records_price():
    item = {'preu': 20.0}
    app._apply_preu_historial(item, None, '2026-08-05')
    assert item['historial'] == [{'preu': 20.0, 'data': '2026-08-05'}]


def test_apply_historial_zero_price_records_nothing():
    item = {'preu': 0}
    app._apply_preu_historial(item, None, '2026-08-05')
    assert item['historial'] == []


def test_apply_historial_unchanged_price_no_new_entry():
    prev = {'preu': 20.0, 'historial': [{'preu': 20.0, 'data': '2026-01-01'}]}
    item = {'preu': 20.0}
    app._apply_preu_historial(item, prev, '2026-08-05')
    assert item['historial'] == [{'preu': 20.0, 'data': '2026-01-01'}]  # sense canvi


def test_apply_historial_changed_price_appends():
    prev = {'preu': 20.0, 'historial': [{'preu': 20.0, 'data': '2026-01-01'}]}
    item = {'preu': 22.5}
    app._apply_preu_historial(item, prev, '2026-08-05')
    assert item['historial'] == [
        {'preu': 20.0, 'data': '2026-01-01'},
        {'preu': 22.5, 'data': '2026-08-05'},
    ]


def test_consum_nou_id_unique_per_index():
    a = app._consum_nou_id(0)
    b = app._consum_nou_id(1)
    assert a != b and a.startswith('c')


# ── Correu en castellà + catàleg Norilab ───────────────────────────────────
def test_email_and_subject_in_spanish():
    grup = {'equip_label': 'Noritsu Green IV', 'proveidor_nom': 'Norilab',
            'items': [{'nom': 'Tinta cyan', 'referencia': 'H086163-00', 'quantitat': 2}]}
    html = app._consumibles_email_html(grup)
    assert 'Buenos días' in html
    assert 'Cantidad' in html and 'Referencia' in html
    assert 'Muchas gracias' in html
    assert app._consumibles_email_subject(grup).startswith('Pedido de consumibles')


def test_norilab_catalog_has_codes_and_prices():
    cat = app.CONSUM_CATALEG_NORILAB
    assert len(cat) == 9
    refs = {c['referencia'] for c in cat}
    assert {'H086162-00', 'H086163-00', 'H086164-00', 'H086165-00'} <= refs
    assert all(c['preu'] > 0 and c['referencia'] and c['nom'] for c in cat)


def test_defaults_include_norilab_catalog():
    refs = {c.get('referencia') for c in app.CONSUMIBLES_DEFAULTS}
    assert 'H086162-00' in refs and 'PAPRL1000128' in refs
    for c in app.CONSUMIBLES_DEFAULTS:
        assert app._normalize_consumible(c)['nom']  # tots normalitzables amb nom


def test_norilab_proveidor_data():
    assert app.CONSUM_PROVEIDOR_NORILAB['email'] == 'administracion@norilabiberia.es'
    assert app.CONSUM_PROVEIDOR_NORILAB['nom']

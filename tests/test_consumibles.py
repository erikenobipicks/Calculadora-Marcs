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
        'proveidor_nom': 'Prov <X>',
        'items': [
            {'nom': 'Cartutx cian', 'referencia': 'ABC-1', 'quantitat': 2, 'equip': 'noritsu_green_iv'},
            {'nom': 'Paper <lustre>', 'referencia': '', 'quantitat': 1, 'equip': 'noritsu_green_iv'},
        ],
    }
    html = app._consumibles_email_html(grup)
    assert 'Equipo' in html                  # columna d'equip
    assert 'Noritsu Green IV' in html        # etiqueta d'equip a la columna
    assert 'Cartutx cian' in html
    assert 'ABC-1' in html
    assert '&lt;lustre&gt;' in html          # nom escapat
    assert '<lustre>' not in html            # cap injecció crua
    assert 'Prov &lt;X&gt;' in html          # proveïdor escapat
    assert '—' in html                       # referència buida → guió


def test_prov_efectiu_override_and_fallback():
    equip_prov = {'canon_pro_4000': {'nom': 'Delex', 'email': 'delex@x.com'}}
    # Override per consumible → mana.
    assert app._consum_prov_efectiu(
        {'equip': 'canon_pro_4000', 'proveidor_nom': 'Norilab', 'proveidor_email': 'n@x.com'},
        equip_prov) == ('Norilab', 'n@x.com')
    # Sense override → proveïdor de l'equip.
    assert app._consum_prov_efectiu(
        {'equip': 'canon_pro_4000', 'proveidor_nom': '', 'proveidor_email': ''},
        equip_prov) == ('Delex', 'delex@x.com')


def test_comanda_grups_by_supplier(monkeypatch):
    items = [
        {'equip': 'noritsu_green_iv', 'nom': 'Cian', 'referencia': '', 'quantitat': 1, 'pendent': True,  'proveidor_nom': '', 'proveidor_email': '', 'notes': '', 'actiu': True, 'ordre': 1},
        {'equip': 'noritsu_green_iv', 'nom': 'Negre', 'referencia': '', 'quantitat': 1, 'pendent': False, 'proveidor_nom': '', 'proveidor_email': '', 'notes': '', 'actiu': True, 'ordre': 2},
        {'equip': 'canon_pro_4000',   'nom': 'Tinta Delex',   'referencia': 'PFI1300C', 'quantitat': 1, 'pendent': True, 'proveidor_nom': '', 'proveidor_email': '', 'notes': '', 'actiu': True, 'ordre': 3},
        {'equip': 'canon_pro_4000',   'nom': 'Tinta Norilab', 'referencia': 'PFI1300C', 'quantitat': 1, 'pendent': True, 'proveidor_nom': 'Norilab', 'proveidor_email': 'norilab@x.com', 'notes': '', 'actiu': True, 'ordre': 4},
    ]
    monkeypatch.setattr(app, 'get_consumibles_list', lambda: items)
    # Tots els equips comparteixen el mateix email de proveïdor 'lab@x.com'.
    monkeypatch.setattr(app, 'get_config_value',
                        lambda clau, default=None: ('lab@x.com' if clau.endswith('_email')
                                                    else 'Lab' if clau.endswith('_nom') else default))
    grups = app._consumibles_comanda_grups()
    emails = sorted(g['proveidor_email'] for g in grups)
    # Els dos sense override (Noritsu Cian + Canon Delex) → mateix proveïdor d'equip → 1 grup.
    # El de Norilab (override) → grup a part. El 'Negre' no és pendent.
    assert emails == ['lab@x.com', 'norilab@x.com']
    lab = next(g for g in grups if g['proveidor_email'] == 'lab@x.com')
    assert len(lab['items']) == 2
    nor = next(g for g in grups if g['proveidor_email'] == 'norilab@x.com')
    assert len(nor['items']) == 1 and nor['proveidor_nom'] == 'Norilab'


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


def test_delex_catalog_and_registry():
    d = app.CONSUM_CATALEG_DELEX
    refs = {c['referencia'] for c in d}
    assert 'PFI1300PBK' in refs and 'T46K140' in refs and 'HQR23061' in refs
    assert all(c['preu'] > 0 and c['referencia'] and c['nom'] for c in d)
    # Tintes Canon → equip Canon; tintes Epson → equip Epson.
    assert all(c['equip'] == 'canon_pro_4000' for c in d if c['referencia'].startswith('PFI'))
    assert all(c['equip'] == 'epson_surelab_d1000' for c in d if c['referencia'].startswith('T46K'))
    # Registre de catàlegs.
    assert {'norilab', 'delex', 'norilab_canon'} <= set(app.CONSUM_CATALEGS)
    assert app.CONSUM_CATALEGS['delex']['proveidor']['email'] == 'info@delex.es'
    assert app.CONSUM_CATALEGS['delex']['equips'] == ['canon_pro_4000', 'epson_surelab_d1000']


def test_defaults_include_both_catalogs():
    refs = {c.get('referencia') for c in app.CONSUMIBLES_DEFAULTS}
    assert {'H086162-00', 'PFI1300PBK', 'T46K140'} <= refs


def test_normalize_includes_proveidor_override():
    c = app._normalize_consumible({'nom': 'X', 'proveidor_nom': ' Norilab ', 'proveidor_email': ' n@x.com '})
    assert c['proveidor_nom'] == 'Norilab'
    assert c['proveidor_email'] == 'n@x.com'


def test_norilab_canon_catalog():
    cat = app.CONSUM_CATALEG_NORILAB_CANON
    assert len(cat) == 12
    assert all(c['preu'] == 144.0 and c['equip'] == 'canon_pro_4000' for c in cat)
    assert all(c['proveidor']['email'] == 'administracion@norilabiberia.es' for c in cat)
    # Registrat i sense fixar proveïdor d'equip (els ítems ja porten proveïdor propi).
    assert 'norilab_canon' in app.CONSUM_CATALEGS
    assert app.CONSUM_CATALEGS['norilab_canon']['equips'] == []


# ── Historial de comandes + tornar a demanar ───────────────────────────────
def test_get_comandes_history_parse(monkeypatch):
    import json as _json
    data = [{'id': 'c1', 'data': '2026-08-05 10:00', 'proveidor_nom': 'Norilab', 'linies': []}]
    monkeypatch.setattr(app, 'get_config_value',
                        lambda clau, default=None: (_json.dumps(data)
                                                    if clau == 'consumibles_comandes_json' else default))
    h = app.get_consumibles_comandes()
    assert len(h) == 1 and h[0]['id'] == 'c1'


def test_get_comandes_history_empty(monkeypatch):
    monkeypatch.setattr(app, 'get_config_value', lambda clau, default=None: default)
    assert app.get_consumibles_comandes() == []


def test_reorder_apply_by_id_and_ref_fallback():
    items = [
        {'id': 'A', 'equip': 'noritsu_green_iv', 'referencia': 'H1', 'quantitat': 1, 'pendent': False},
        {'id': 'B', 'equip': 'canon_pro_4000', 'referencia': 'PFI1300C', 'quantitat': 1, 'pendent': False},
    ]
    linies = [
        {'consumible_id': 'A', 'referencia': 'H1', 'equip': 'noritsu_green_iv', 'quantitat': 3},        # per id
        {'consumible_id': 'X', 'referencia': 'PFI1300C', 'equip': 'canon_pro_4000', 'quantitat': 2},    # id perdut → ref+equip
        {'consumible_id': 'Y', 'referencia': 'NOEXIST', 'equip': 'canon_pro_4000', 'quantitat': 1},     # no existeix
    ]
    marcats, no_trobats = app._consumibles_reorder_apply(items, linies)
    assert marcats == 2 and no_trobats == 1
    assert items[0]['pendent'] is True and items[0]['quantitat'] == 3
    assert items[1]['pendent'] is True and items[1]['quantitat'] == 2


def test_reorder_apply_empty():
    assert app._consumibles_reorder_apply([], None) == (0, 0)

"""Tests del generador de tarifes (_tarifa_collect_data, _tarifa_make_ref...).

L'autouse fixture de conftest.py monkeypatcha `app.query` a `[]` i
`app.get_config_value` a tornar el default, així `_imp_closest` cau a
formula i `_tarifa_default_sizes` retorna sempre [] (cap fila a la BD).

Aquests tests bloquegen la regressió descrita al reporte de bugs:
referències duplicades al xlsx d'export per a mides addicionals que
internament mapegen a una fila estàndard via min-contain.
"""
import pytest

import app


def test_impressio_custom_sizes_get_unique_refs(monkeypatch):
    """Dues mides addicionals diferents que internament cauen al mateix
    ref de catàleg han de rebre refs úniques a l'export."""

    # Simulem que _imp_closest retorna el MATEIX ref de catàleg per a dues
    # mides diferents (reproduint el cas del bug: 9×13 i 10×15 → 'IMP10x15').
    def fake_imp_closest(w, h, paper='lustre'):
        if (w, h) in [(9, 13), (10, 15)]:
            return {'ref': 'IMP10x15', 'preu': 2.50, 'origen': 'taula', 'area': w * h}
        return {'ref': f'imp-{w}x{h}', 'preu': 1.0, 'origen': 'formula', 'area': w * h}

    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    custom_sizes = {'impressio': [(9, 13), (10, 15)]}
    data = app._tarifa_collect_data(['impressio'], custom_sizes, usuari=None)

    assert len(data) == 1 and data[0]['key'] == 'impressio'
    rows = data[0]['rows']
    refs = [r['ref'] for r in rows]
    # Cada (w, h) ha de tenir una ref única.
    assert len(refs) == len(set(refs)), f'refs duplicades: {refs}'
    # I el format ha de ser IMP{w}x{h} (uniforme amb el catàleg).
    by_size = {(r['w'], r['h']): r['ref'] for r in rows}
    assert by_size[(9, 13)] == 'IMP9x13'
    assert by_size[(10, 15)] == 'IMP10x15'


def test_impressio_formula_path_also_normalized(monkeypatch):
    """Mides addicionals que internament cauen a fórmula (ref legacy
    'imp-{w}x{h}' en minúscules) també han de quedar normalitzades a
    'IMP{w}x{h}' a l'export per consistència de format."""

    def fake_imp_closest(w, h, paper='lustre'):
        return {'ref': f'imp-{w}x{h}', 'preu': 5.0, 'origen': 'formula', 'area': w * h}
    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    custom_sizes = {'impressio': [(25, 25), (60, 60)]}
    data = app._tarifa_collect_data(['impressio'], custom_sizes, usuari=None)
    refs = [r['ref'] for r in data[0]['rows']]
    assert 'IMP25x25' in refs
    assert 'IMP60x60' in refs
    # Cap ref en format legacy minúscula-amb-guió.
    assert not any(r.startswith('imp-') for r in refs)


def test_impressio_size_in_catalog_keeps_catalog_ref(monkeypatch):
    """Si una mida que l'usuari passa com a 'addicional' coincideix
    exactament amb una fila de catàleg, l'export ha de mantenir la
    ref del catàleg (no sobreescriure amb una de generada)."""

    # Cataleg amb IMP10x15 (només).
    def fake_query(sql, args=()):
        if 'FROM impressio' in sql:
            return [{'referencia': 'IMP10x15'}]
        return []
    monkeypatch.setattr(app, 'query', fake_query)

    def fake_imp_closest(w, h, paper='lustre'):
        if (w, h) == (10, 15):
            return {'ref': 'IMP10x15', 'preu': 2.50, 'origen': 'taula', 'area': 150}
        return None
    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    # L'usuari passa 10×15 com a "custom" — ja és catàleg.
    custom_sizes = {'impressio': [(10, 15)]}
    data = app._tarifa_collect_data(['impressio'], custom_sizes, usuari=None)
    rows = data[0]['rows']
    # No duplicació: la mida apareix una sola vegada, amb ref del catàleg.
    assert len(rows) == 1
    assert rows[0]['ref'] == 'IMP10x15'


def test_impressio_collision_with_catalog_ref_is_resolved(monkeypatch):
    """Cas del bug reportat: la mida addicional 40×40 mapeja
    internament a 'IMP40x50' (catàleg) via _imp_closest, però l'export
    ha de donar-li IMP40x40 perquè no es solapi amb la fila real."""

    def fake_query(sql, args=()):
        if 'FROM impressio' in sql:
            # Catàleg conté IMP40x50 (estàndard) però NO IMP40x40.
            return [{'referencia': 'IMP40x50'}]
        return []
    monkeypatch.setattr(app, 'query', fake_query)

    def fake_imp_closest(w, h, paper='lustre'):
        # _imp_closest mapeja 40×40 → IMP40x50 (taula, ratio dins threshold).
        # 40×50 → IMP40x50 directament.
        if (w, h) in [(40, 40), (40, 50)]:
            return {'ref': 'IMP40x50', 'preu': 9.50, 'origen': 'taula', 'area': w * h}
        return None
    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    # 40×40 és custom. 40×50 és del catàleg (apareix per _tarifa_default_sizes).
    custom_sizes = {'impressio': [(40, 40)]}
    data = app._tarifa_collect_data(['impressio'], custom_sizes, usuari=None)
    rows = data[0]['rows']
    refs_by_size = {(r['w'], r['h']): r['ref'] for r in rows}
    assert refs_by_size[(40, 40)] == 'IMP40x40', 'custom size hauria de generar ref pròpia'
    assert refs_by_size[(40, 50)] == 'IMP40x50', 'catàleg conserva la seva ref'
    # Cap col·lisió.
    refs = list(refs_by_size.values())
    assert len(refs) == len(set(refs))

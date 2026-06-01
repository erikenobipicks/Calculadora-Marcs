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


# ── Option B: max(small_calib, big_calib) a _imp_closest ──────────────────


def test_imp_closest_formula_uses_max_of_two_calibrations(monkeypatch):
    """A la branca fórmula, _imp_closest ha d'avaluar dues calibracions
    (ref menor o igual i ref major o igual) i tornar el MAX de les dues.
    Sense això una mida intermèdia subestima el preu (bug 2 del reporte
    d'impressions: 60×60 sortia 10.80 quan hauria d'estar prop de 14)."""
    # Stub de query: catàleg amb una ref petita barata i una gran cara,
    # cap dins ratio_max per a 60×60 (3600 cm²).
    def fake_query(sql, args=()):
        if 'FROM impressio' in sql:
            # IMP50x70 = 3500 cm² a preu 10€ (€/cm² = 0.00286)
            # IMP60x100 = 6000 cm² a preu 24€ (€/cm² = 0.004)
            return [
                {'referencia': 'IMP50x70',  'preu': 10.0},
                {'referencia': 'IMP60x100', 'preu': 24.0},
            ]
        return []
    monkeypatch.setattr(app, 'query', fake_query)

    r = app._imp_closest(60, 60)
    assert r is not None
    # IMP60x100 (6000 cm²) → ratio 6000/3600 = 1.67 > 1.40 → fora threshold
    # → ha de caure a fórmula. Calibració amb IMP60x100 dóna 3600/6000 ·
    # 24€ = 14.40€. Calibració amb IMP50x70 dóna ~ 10.28€. Max = 14.40€.
    assert r['origen'] == 'formula'
    assert r['preu'] >= 14.0, f'Option B hauria de donar preu prop de 14.40; got {r["preu"]}'


# ── Floor: catàleg × factor com a sòl per a totes les mides ───────────────


def test_floor_applies_when_size_exists_in_catalog(monkeypatch):
    """Si una mida custom coincideix exactament amb una fila de catàleg,
    el PVD final no pot baixar per sota de catalog_pvd × factor_floor.
    Default factor = 1.03 (3% mínim per sobre del catàleg actual)."""
    # Catàleg: 60×60 a 14.50€. Si _imp_closest retorna 10.80 (formula
    # path), el floor ha de pujar el PVD a 14.50 · 1.03 = 14.94.
    def fake_query(sql, args=()):
        if 'FROM impressio' in sql:
            return [{'referencia': 'IMP60x60', 'preu': 14.50}]
        return []
    monkeypatch.setattr(app, 'query', fake_query)

    # Simulem _imp_closest tornant un preu baix (formula via mock).
    def fake_imp_closest(w, h, paper='lustre'):
        return {'ref': 'IMP60x60', 'preu': 10.80, 'origen': 'formula', 'area': 3600}
    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    data = app._tarifa_collect_data(['impressio'], {'impressio': []}, usuari=None)
    rows = data[0]['rows']
    by_size = {(r['w'], r['h']): r for r in rows}
    assert (60, 60) in by_size
    assert by_size[(60, 60)]['pvd'] == round(14.50 * 1.03, 2), \
        f'floor 14.50·1.03 hauria d\'aplicar-se; got {by_size[(60, 60)]["pvd"]}'
    assert by_size[(60, 60)]['origen'] == 'floor'


def test_floor_does_not_lower_already_higher_price(monkeypatch):
    """Si el càlcul ja dóna més que el floor, no es toca."""
    def fake_query(sql, args=()):
        if 'FROM impressio' in sql:
            return [{'referencia': 'IMP60x60', 'preu': 14.50}]
        return []
    monkeypatch.setattr(app, 'query', fake_query)

    # Càlcul base = 20€ (per sobre del floor 14.94).
    def fake_imp_closest(w, h, paper='lustre'):
        return {'ref': 'IMP60x60', 'preu': 20.0, 'origen': 'taula', 'area': 3600}
    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    data = app._tarifa_collect_data(['impressio'], {'impressio': []}, usuari=None)
    pvd = data[0]['rows'][0]['pvd']
    assert pvd == 20.0, 'preus superiors al floor no s\'han de modificar'


def test_floor_skipped_for_sizes_not_in_catalog(monkeypatch):
    """Per a mides que NO estan al catàleg, el floor no s'aplica (no hi
    ha referència històrica per comparar)."""
    def fake_query(sql, args=()):
        if 'FROM impressio' in sql:
            return [{'referencia': 'IMP60x60', 'preu': 14.50}]
        return []
    monkeypatch.setattr(app, 'query', fake_query)

    def fake_imp_closest(w, h, paper='lustre'):
        return {'ref': f'imp-{w}x{h}', 'preu': 3.0, 'origen': 'formula', 'area': w * h}
    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    custom = {'impressio': [(25, 25)]}  # NO al catàleg.
    data = app._tarifa_collect_data(['impressio'], custom, usuari=None)
    by_size = {(r['w'], r['h']): r for r in data[0]['rows']}
    # 25×25 no està al catàleg → cap floor, pvd queda al càlcul base.
    assert by_size[(25, 25)]['pvd'] == 3.0
    # 60×60 sí està al catàleg i el càlcul (3€) està per sota → floor.
    assert by_size[(60, 60)]['pvd'] == round(14.50 * 1.03, 2)


def test_floor_factor_configurable(monkeypatch):
    """El factor del floor es llegeix de la clau 'tarifa_floor_factor'."""
    def fake_query(sql, args=()):
        if 'FROM impressio' in sql:
            return [{'referencia': 'IMP60x60', 'preu': 14.50}]
        return []
    monkeypatch.setattr(app, 'query', fake_query)

    def fake_get_config_value(clau, default=None):
        if clau == 'tarifa_floor_factor':
            return '1.10'  # 10% mínim
        return default
    monkeypatch.setattr(app, 'get_config_value', fake_get_config_value)

    def fake_imp_closest(w, h, paper='lustre'):
        return {'ref': 'IMP60x60', 'preu': 10.0, 'origen': 'formula', 'area': 3600}
    monkeypatch.setattr(app, '_imp_closest', fake_imp_closest)

    data = app._tarifa_collect_data(['impressio'], {'impressio': []}, usuari=None)
    pvd = data[0]['rows'][0]['pvd']
    assert pvd == round(14.50 * 1.10, 2), f'factor 1.10 hauria de donar 15.95; got {pvd}'


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

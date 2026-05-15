"""Regression tests per a les funcions calcular_cost_* / calcular_pvd / calcular_preu_marc.

Estratègia:
- Cap accés a BD: `query` retorna [] (fixture autouse a conftest.py).
- Cap valor de config llegit: `get_config_value` retorna el `default` que
  la pròpia funció passa → els tests fan servir EXACTAMENT els defaults
  que veu producció en una BD recent seedada.
- Tot el flux cau a la branca FÓRMULA, que és matemàticament determinista.
- Valors esperats calculats a mà a partir de les fórmules + defaults.

Aquests tests NO validen "el preu és correcte de negoci" — només "la
fórmula actual no s'ha mogut sense que ningú se n'adoni". Si algú toca
una fórmula o un default seed, el test cau i obliga a justificar-ho.
"""
import math
import pytest

import app


APPROX = lambda v: pytest.approx(v, abs=1e-4)


# ── calcular_pvd ──────────────────────────────────────────────────────

def test_pvd_aplica_marge_60_per_defecte():
    assert app.calcular_pvd(10.0, 'moldures') == APPROX(16.0)
    assert app.calcular_pvd(10.0, 'vidres') == APPROX(16.0)
    assert app.calcular_pvd(10.0, 'passpartu') == APPROX(16.0)
    assert app.calcular_pvd(10.0, 'encolat') == APPROX(16.0)


def test_pvd_zero_i_none():
    assert app.calcular_pvd(0, 'moldures') == APPROX(0.0)
    assert app.calcular_pvd(None, 'moldures') is None


# ── calcular_preu_marc ────────────────────────────────────────────────

def test_preu_marc_30x40_gruix_2_preu_0_50():
    # perim=140, long_bruta=(140+16)*1.10=171.6, >100 → long=171.6
    # cost = 171.6 * 0.50 / 100 = 0.858; pvd = 0.858 * 1.60
    r = app.calcular_preu_marc(30, 40, 2.0, 0.50)
    assert r['cost'] == APPROX(0.858)
    assert r['pvd'] == APPROX(1.3728)


def test_preu_marc_minim_cm_aplica_quan_motllura_petita():
    # 10x10, gruix 1, preu 1.0: long_bruta = (40+8)*1.10 = 52.8 → < 100 → 100
    r = app.calcular_preu_marc(10, 10, 1.0, 1.0)
    assert r['cost'] == APPROX(1.0)
    assert r['pvd'] == APPROX(1.6)


def test_preu_marc_retorna_none_si_falten_dades():
    assert app.calcular_preu_marc(0, 40, 2, 0.50) is None
    assert app.calcular_preu_marc(30, 40, 2, None) is None


# ── calcular_cost_passpartu — simple, fórmula ─────────────────────────

@pytest.mark.parametrize('w,h,cost_esperat', [
    (30, 40, 4.494),    # mat=max(0.744,0.5)=0.744 ; mo=3.75 → 4.494
    (50, 70, 5.92),     # mat=2.17 ; mo=3.75 → 5.92
    (80, 100, 8.71),    # mat=4.96 ; mo=3.75 → 8.71
])
def test_passpartu_simple_formula(w, h, cost_esperat):
    r = app.calcular_cost_passpartu(w, h, tipus='simple')
    assert r['cost'] == APPROX(cost_esperat)
    assert r['pvd'] == APPROX(cost_esperat * 1.60)
    assert r['origen'] == 'formula'
    assert r['ref'] == f'pas-{w}x{h}'


def test_passpartu_minim_material_aplica_quan_area_petita():
    # 10x10: area=100, mat brut = 100*0.000620 = 0.062 → minim_material=0.50 → 0.50
    # cost = 0.50 + 3.75 = 4.25
    r = app.calcular_cost_passpartu(10, 10, tipus='simple')
    assert r['cost'] == APPROX(4.25)


def test_passpartu_simple_amb_finestres_extra():
    # 50x70 + 2 finestres: 5.92 + 2 * (3.5*25/60) = 5.92 + 2.9167 = 8.8367
    r = app.calcular_cost_passpartu(50, 70, tipus='simple', finestres_extra=2)
    assert r['cost'] == APPROX(8.8367)


# ── calcular_cost_passpartu — doble, simple_x2 ────────────────────────

def test_passpartu_doble_es_simple_x2_quan_no_hi_ha_taula():
    # Sense taula, doble cau a simple_x2 sempre.
    simple = app.calcular_cost_passpartu(50, 70, tipus='simple')
    doble = app.calcular_cost_passpartu(50, 70, tipus='doble')
    assert doble['cost'] == APPROX(simple['cost'] * 2)
    assert doble['origen'] == 'simple_x2'
    assert doble['ref'].startswith('dobpas-')


def test_passpartu_doble_amb_finestres_afegeix_temps_extra():
    # doble 50x70 + 2 finestres:
    #   simple_x2 = 5.92*2 = 11.84
    #   extra = 2 * (3.5*25/60) = 2.9167
    #   total = 14.7567
    r = app.calcular_cost_passpartu(50, 70, tipus='doble', finestres_extra=2)
    assert r['cost'] == APPROX(14.7567)


# ── calcular_cost_foam — fórmula ──────────────────────────────────────

@pytest.mark.parametrize('w,h,cost_esperat', [
    # mat = area*0.001143 ; temps = 9 + area*0.0015 ; mo = temps*25/60
    (30, 40, 5.8716),   # area=1200 → mat=1.3716, temps=10.8, mo=4.5
    (50, 70, 9.938),    # area=3500 → mat=4.0005, temps=14.25, mo=5.9375
    (80, 100, 17.894),  # area=8000 → mat=9.144, temps=21, mo=8.75
])
def test_foam_formula(w, h, cost_esperat):
    r = app.calcular_cost_foam(w, h)
    assert r['origen'] == 'formula'
    assert r['ref'] == f'foam-{w}x{h}'
    # Comprovació estructural (pvd = cost * 1.60)
    assert r['pvd'] == APPROX(r['cost'] * 1.60)
    assert r['preu'] == r['pvd']
    # Comprovació exacta del cost (regression):
    assert r['cost'] == APPROX(cost_esperat)


# ── calcular_cost_laminat ─────────────────────────────────────────────

@pytest.mark.parametrize('w,h,tipus,cost_esperat', [
    # semi: mat=area*0.000504 ; temps=12+area*0.0012 ; mo=temps*25/60
    (30, 40, 'semibrillo', 6.2048),
    (50, 70, 'semibrillo', 8.514),
    # mate: mat=area*0.000685 (mateix temps)
    (30, 40, 'mate', 6.422),
    (50, 70, 'mate', 9.1475),
])
def test_laminat_formula(w, h, tipus, cost_esperat):
    r = app.calcular_cost_laminat(w, h, tipus=tipus)
    assert r['cost'] == APPROX(cost_esperat)
    assert r['origen'] == 'formula'
    assert r['tipus'] == tipus
    assert r['ref'] == f'laminat-{tipus}-{w}x{h}'


# ── calcular_cost_protter — composició foam + laminat ────────────────

@pytest.mark.parametrize('tipus', ['semibrillo', 'mate'])
def test_protter_es_foam_mes_laminat(tipus):
    foam = app.calcular_cost_foam(50, 70)['cost']
    lam = app.calcular_cost_laminat(50, 70, tipus=tipus)['cost']
    prot = app.calcular_cost_protter(50, 70, tipus=tipus)
    assert prot['cost'] == APPROX(round(foam + lam, 4))
    assert prot['origen'] == 'composicio'
    assert prot['tipus'] == tipus


def test_protter_30x40_semibrillo_valor_exacte():
    # foam 5.8716 + laminat semi 6.2048 = 12.0764
    r = app.calcular_cost_protter(30, 40, tipus='semibrillo')
    assert r['cost'] == APPROX(12.0764)


# ── calcular_cost_vidre — fórmula (sense taula) ──────────────────────

@pytest.mark.parametrize('w,h,cost_esperat', [
    # mat = area*0.002880 ; temps = 3 + 0.5 * perim_m ; mo = temps*25/60
    (30, 40, 4.9977),   # area=1200, perim_m=1.4, temps=3.7, mo=1.5417 → cost≈4.9977
    (50, 70, 11.83),    # area=3500, perim_m=2.4, temps=4.2, mo=1.75 → 11.83
    (80, 100, 25.04),   # area=8000, perim_m=3.6, temps=4.8, mo=2.0 → 25.04
])
def test_vidre_simple_formula(w, h, cost_esperat):
    r = app.calcular_cost_vidre(w, h)
    assert r['cost'] == APPROX(cost_esperat)
    assert r['origen'] == 'formula'
    assert r['ref'] == f'vidre-{w}x{h}'
    assert r['pvd'] == APPROX(round(r['cost'] * 1.60, 4))


# ── calcular_cost_doble_vidre — fórmula pura ─────────────────────────

@pytest.mark.parametrize('w,h,cost_esperat', [
    # mat = area*0.002880*2 ; mo_tall = (3 + 0.5*perim_m)*25/60 ; muntatge=1.30
    (30, 40, 9.7537),   # mat=6.912, perim_m=1.4, mo_tall=1.5417 → 9.7537
    (50, 70, 23.21),    # mat=20.16, perim_m=2.4, mo_tall=1.75 → 23.21
    (80, 100, 49.38),   # mat=46.08, perim_m=3.6, mo_tall=2.0 → 49.38
])
def test_doble_vidre_formula(w, h, cost_esperat):
    r = app.calcular_cost_doble_vidre(w, h)
    assert r['cost'] == APPROX(cost_esperat)
    assert r['origen'] == 'formula'
    assert r['ref'] == f'dv-{w}x{h}'


# ── calcular_cost_mirall — facturat per múltiples de 6 dm² ──────────

@pytest.mark.parametrize('w,h,area_fact_dm2,cost_esperat', [
    # area_real_dm2 = w*h/100 ; area_fact_dm2 = ceil(area/6)*6
    # cost = area_fact_cm2 * 0.003153
    (30, 40, 12, 3.7836),   # 12 dm² real → 12 dm² fact (12/6=2 exacte)
    (50, 70, 36, 11.3508),  # 35 → ceil(35/6)*6 = 6*6 = 36
    (80, 100, 84, 26.4852), # 80 → ceil(80/6)*6 = 14*6 = 84
])
def test_mirall_formula(w, h, area_fact_dm2, cost_esperat):
    r = app.calcular_cost_mirall(w, h)
    assert r['cost'] == APPROX(cost_esperat)
    assert r['origen'] == 'formula'
    assert r['area_fact_cm2'] == APPROX(area_fact_dm2 * 100)
    assert r['ref'] == f'mir-{w}x{h}'


# ── Invariants estructurals (no haurien de canviar mai) ───────────────

def test_totes_funcions_retornen_dict_amb_cost_i_pvd():
    funcions = [
        ('passpartu_simple', lambda: app.calcular_cost_passpartu(50, 70, 'simple')),
        ('passpartu_doble',  lambda: app.calcular_cost_passpartu(50, 70, 'doble')),
        ('foam',             lambda: app.calcular_cost_foam(50, 70)),
        ('laminat_semi',     lambda: app.calcular_cost_laminat(50, 70, 'semibrillo')),
        ('laminat_mate',     lambda: app.calcular_cost_laminat(50, 70, 'mate')),
        ('protter_semi',     lambda: app.calcular_cost_protter(50, 70, 'semibrillo')),
        ('vidre',            lambda: app.calcular_cost_vidre(50, 70)),
        ('doble_vidre',      lambda: app.calcular_cost_doble_vidre(50, 70)),
        ('mirall',           lambda: app.calcular_cost_mirall(50, 70)),
    ]
    for nom, fn in funcions:
        r = fn()
        assert isinstance(r, dict), f'{nom} no retorna dict'
        assert 'cost' in r and r['cost'] is not None, f'{nom} sense cost'
        assert 'pvd'  in r and r['pvd']  is not None, f'{nom} sense pvd'
        assert 'origen' in r, f'{nom} sense origen'
        assert 'ref'    in r, f'{nom} sense ref'
        assert r['cost'] > 0, f'{nom} cost <= 0'
        # PVD ha de ser exactament cost · 1.60 amb defaults
        assert r['pvd'] == APPROX(round(r['cost'] * 1.60, 4)), f'{nom} pvd != cost*1.6'

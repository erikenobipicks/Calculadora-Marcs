"""Factura Directa des de la cistella ha d'aplicar els descomptes.

Bug: `_linies_de_cistella` construïa les línies FD amb el preu brut, ignorant
el descompte per línia i el descompte global del pressupost. Resultat: el
document de FacturaDirecta sobrefacturava respecte de la cistella a pantalla,
el PDF (crear_pdf_marcs) i el desat (api_desar_cistella), que sí els apliquen.
"""
import app


def _unit(items, mode='pvp', dg=0):
    linies = app._linies_de_cistella(items, mode, recarrec=False, descompte_global=dg)
    return [l['unitPrice'] for l in linies]


def test_sense_descompte():
    assert _unit([{'text': 'X', 'quantity': 1, 'preu_net': 100.0}]) == [100.0]


def test_descompte_per_linia():
    assert _unit([{'text': 'X', 'quantity': 1, 'preu_net': 100.0, 'descompte': 10}]) == [90.0]


def test_descompte_global():
    assert _unit([{'text': 'X', 'quantity': 1, 'preu_net': 100.0}], dg=20) == [80.0]


def test_descompte_linia_i_global_es_combinen():
    # eff = 1-(1-0.10)*(1-0.10) = 0.19 → 100 * 0.81 = 81.00
    assert _unit([{'text': 'X', 'quantity': 1, 'preu_net': 100.0, 'descompte': 10}], dg=10) == [81.0]


def test_reparteix_per_unitats():
    # 5 u, preu_net total 50, descompte 10% → net 45 / 5 = 9,00 unitat
    assert _unit([{'text': 'X', 'quantity': 5, 'preu_net': 50.0, 'descompte': 10}]) == [9.0]


def test_mode_cost_usa_cost_produccio():
    items = [{'text': 'X', 'quantity': 1, 'preu_net': 100.0, 'cost_produccio': 60.0, 'descompte': 10}]
    assert _unit(items, mode='cost') == [54.0]  # 60 * 0.90

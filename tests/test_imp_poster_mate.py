"""El Pòster Mate no ha de sortir mai més car que el Lustre.

Bug: el Mate només té tarifa de gran format (anca mínima 40x50 = 8 €). Per a
mides petites, el blindatge d'extrapolació el clavava a 8 € mentre que el
Lustre baixava fins a <1 €, així que el Mate —un paper més econòmic— sortia
més car. `_imp_closest` ara cap el Mate a una fracció del preu Lustre.
"""
import app

# Catàleg sintètic: anques de gran format del Mate (prefix MATE-) + Lustre
# (sense prefix) amb mides petites i grans, coherent amb producció.
_ROWS = [
    # Lustre (sense prefix): mides petites + alguna de gran format
    {'referencia': '20X15', 'descripcio': '', 'preu': 0.87},
    {'referencia': '30X20', 'descripcio': '', 'preu': 1.79},
    {'referencia': '30X40', 'descripcio': '', 'preu': 3.60},
    {'referencia': '40X50', 'descripcio': '', 'preu': 10.29},
    {'referencia': '50X70', 'descripcio': '', 'preu': 15.88},
    {'referencia': '100X150', 'descripcio': '', 'preu': 66.21},
    # Pòster Mate (prefix MATE-): NOMÉS gran format
    {'referencia': 'MATE-40X50', 'descripcio': '', 'preu': 8.00},
    {'referencia': 'MATE-50X70', 'descripcio': '', 'preu': 11.76},
    {'referencia': 'MATE-100X150', 'descripcio': '', 'preu': 45.52},
]


def _use_rows(monkeypatch):
    monkeypatch.setattr(app, 'query',
                        lambda *a, **k: _ROWS if ('impressio' in (a[0] if a else '')) else [])
    # get_config_value ja retorna el default (ratio 0.80) via conftest.


def test_mate_sempre_mes_barat_que_lustre(monkeypatch):
    _use_rows(monkeypatch)
    for (w, h) in [(20, 15), (30, 20), (30, 40), (40, 50), (50, 70), (60, 80), (100, 150)]:
        lp = app._imp_closest(w, h, paper='lustre')['preu']
        mp = app._imp_closest(w, h, paper='poster_mate')['preu']
        assert mp < lp, f"{w}x{h}: Mate {mp} no és més barat que Lustre {lp}"


def test_mate_mides_petites_ja_no_es_claven(monkeypatch):
    _use_rows(monkeypatch)
    # 30x20: abans es clavava a 8,00 €; ara ha de ser <= 80% del Lustre (1,79).
    mp = app._imp_closest(30, 20, paper='poster_mate')['preu']
    assert mp <= round(1.79 * 0.80, 2) + 0.001
    assert mp < 8.0


def test_mate_gran_format_conserva_la_seva_anca(monkeypatch):
    _use_rows(monkeypatch)
    # A gran format el Mate ja és més barat que el 80% del Lustre → es manté l'anca.
    assert app._imp_closest(40, 50, paper='poster_mate')['preu'] == 8.00
    assert app._imp_closest(100, 150, paper='poster_mate')['preu'] == 45.52

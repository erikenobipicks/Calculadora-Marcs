"""Blindatge de `_imp_closest` contra files "verí" a la taula `impressio`.

Bug reportat: una còpia 50x50 sortia a 0,93 € (PVD) mentre que una 40x60
(àrea gairebé igual) sortia a 7,80 €. Causa: hi havia una fila mal
introduïda al catàleg (referència 50X50 amb preu ~0,93 €). El motor de
preus calibra amb la fila més propera per sota i per sobre de l'àrea
sol·licitada i en pren el MAX; això rescata els veïns d'una fila dolenta,
però quan la mida demanada coincideix EXACTAMENT amb l'àrea de la fila
verí, les dues calibracions hi col·lapsen i el MAX no la pot rescatar.

El blindatge: si el preu principal cau per sota del 50% del que prediuen
els veïns ESTRICTES (àrea < i > a la demanada), s'usa la tendència dels
veïns. Aquests tests fixen la regressió i comproven que cap preu legítim
canvia.
"""
import app


# Catàleg Lustre sintètic: anques petites + gran format coherents
# (~0,003-0,005 €/cm²) MÉS una fila verí 50X50 a 0,93 €.
_ROWS = [
    {'referencia': '20X15', 'descripcio': '', 'preu': 0.87},
    {'referencia': '30X20', 'descripcio': '', 'preu': 1.79},
    {'referencia': '40X30', 'descripcio': '', 'preu': 3.60},
    {'referencia': '40X60', 'descripcio': '', 'preu': 7.80},
    {'referencia': '40X50', 'descripcio': '', 'preu': 10.29},
    {'referencia': '50X70', 'descripcio': '', 'preu': 15.88},
    {'referencia': '60X80', 'descripcio': '', 'preu': 21.75},
    {'referencia': '50X50', 'descripcio': '', 'preu': 0.93},  # verí
]


def _use_rows(monkeypatch, rows):
    monkeypatch.setattr(app, 'query',
                        lambda *a, **k: rows if ('impressio' in (a[0] if a else '')) else [])
    # get_config_value ja retorna el default via conftest → trams off, cost_cm2 default.


def test_fila_veri_no_ensorra_el_preu(monkeypatch):
    """La 50x50, que coincideix amb la fila verí, no pot sortir a 0,93 €."""
    _use_rows(monkeypatch, _ROWS)
    r = app._imp_closest(50, 50, paper='lustre')
    # Ha de seguir la tendència dels veïns (40x60=7,80 · 50x70=15,88), no la fila verí.
    assert r['preu'] > 7.0, f"50x50 sortiria a {r['preu']} € (fila verí no blindada)"
    assert r['preu'] < 16.0


def test_preus_legitims_no_canvien(monkeypatch):
    """El blindatge NOMÉS rescata l'outlier; la resta de mides queden igual."""
    _use_rows(monkeypatch, _ROWS)
    esperat = {
        (20, 15): 0.87,
        (30, 20): 1.79,
        (40, 30): 3.60,
        (40, 60): 7.80,
        (40, 50): 10.29,
        (50, 70): 15.88,
    }
    for (w, h), preu in esperat.items():
        got = app._imp_closest(w, h, paper='lustre')['preu']
        assert abs(got - preu) < 0.01, f"{w}x{h}: esperat {preu}, obtingut {got}"


def test_preu_coherent_amb_o_sense_fila_veri(monkeypatch):
    """Amb la fila verí present o corregida, la 50x50 dóna el mateix preu sa."""
    _use_rows(monkeypatch, _ROWS)
    amb = app._imp_closest(50, 50, paper='lustre')['preu']
    _use_rows(monkeypatch, [r for r in _ROWS if r['referencia'] != '50X50'])
    sense = app._imp_closest(50, 50, paper='lustre')['preu']
    assert abs(amb - sense) < 0.01, f"amb={amb} sense={sense}"

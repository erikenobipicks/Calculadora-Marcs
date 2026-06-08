"""Tests del control d'stock de marcs.

- `_consum_marc_cm`: rèplica exacta de la fórmula de consum de la calculadora
  (perímetre exterior + gruix·8, merma i mínim). Determinista.
- `_aplica_moviment_stock`: moviments entrada/sortida/ajust, stock resultant,
  detecció de sota-mínim, opt-in (NULL = no controlat) i activació.

Es fa servir monkeypatch sobre `app.query` / `app.execute` per simular la BD
sense tocar cap motor real.
"""
import pytest

import app


APPROX = lambda v: pytest.approx(v, abs=1e-4)


# ── _consum_marc_cm ───────────────────────────────────────────────────

def test_consum_marc_basic_30x40():
    # per=2*(30+40)=140; (140 + 2*8)*1.10 = 156*1.10 = 171.6; >100
    assert app._consum_marc_cm(30, 40, 2.0, 10.0, 100.0) == APPROX(171.6)


def test_consum_marc_aplica_minim():
    # peça petita: (40 + 8)*1.10 = 52.8 → puja al mínim 100
    assert app._consum_marc_cm(10, 10, 1.0, 10.0, 100.0) == APPROX(100.0)


def test_consum_marc_premarc_suma_gruixos():
    # premarc: gruix_pre=1 + gruix_marc(extra)=2 → (140 + 3*8)*1.10 = 164*1.10 = 180.4
    assert app._consum_marc_cm(30, 40, 1.0, 10.0, 100.0, gruix_extra=2.0) == APPROX(180.4)


def test_consum_marc_dimensions_zero():
    assert app._consum_marc_cm(0, 40, 2.0) == 0.0
    assert app._consum_marc_cm(30, 0, 2.0) == 0.0


# ── _aplica_moviment_stock ────────────────────────────────────────────

def _patch_db(monkeypatch, stock_cm, stock_min_cm):
    """Simula una fila de moldures i captura els execute()."""
    calls = []
    row = {'stock_cm': stock_cm, 'stock_min_cm': stock_min_cm}

    def fake_query(sql, args=(), one=False):
        return row if one else [row]

    def fake_execute(sql, args=()):
        calls.append((sql, list(args)))

    monkeypatch.setattr(app, 'query', fake_query)
    monkeypatch.setattr(app, 'execute', fake_execute)
    return calls


def test_sortida_descompta(monkeypatch):
    calls = _patch_db(monkeypatch, 500.0, 100.0)
    res = app._aplica_moviment_stock('M1', 150, 'sortida')
    assert res['stock_resultant'] == APPROX(350.0)
    assert res['cm'] == APPROX(-150.0)
    assert res['sota_minim'] is False
    # UPDATE + INSERT
    assert len(calls) == 2


def test_entrada_suma(monkeypatch):
    _patch_db(monkeypatch, 500.0, 100.0)
    res = app._aplica_moviment_stock('M1', 100, 'entrada')
    assert res['stock_resultant'] == APPROX(600.0)
    assert res['cm'] == APPROX(100.0)


def test_ajust_fixa_valor_i_detecta_sota_minim(monkeypatch):
    _patch_db(monkeypatch, 500.0, 100.0)
    res = app._aplica_moviment_stock('M1', 80, 'ajust')
    assert res['stock_resultant'] == APPROX(80.0)
    assert res['sota_minim'] is True


def test_optin_null_no_descompta(monkeypatch):
    _patch_db(monkeypatch, None, None)
    assert app._aplica_moviment_stock('M1', 150, 'sortida') is None


def test_activar_inicialitza_des_de_null(monkeypatch):
    _patch_db(monkeypatch, None, None)
    res = app._aplica_moviment_stock('M1', 50, 'ajust', activar=True)
    assert res is not None
    assert res['stock_resultant'] == APPROX(50.0)


def test_ref_buida_o_guio(monkeypatch):
    _patch_db(monkeypatch, 500.0, 100.0)
    assert app._aplica_moviment_stock('', 10, 'entrada') is None
    assert app._aplica_moviment_stock('-', 10, 'entrada') is None

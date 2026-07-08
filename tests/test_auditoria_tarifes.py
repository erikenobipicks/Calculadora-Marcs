"""Tests de l'auditoria de tarifes (_auditoria_tarifes i helpers).

Es simula la BD amb monkeypatch sobre `app.query`, dispatchant per taula
segons l'SQL. Comprovem que cada regla d'anomalia dispara (o no) com toca.
"""
import pytest

import app


def _run(monkeypatch, moldures=None, vidres=None, passpartout=None, encolat_pro=None):
    data = {
        'moldures': moldures or [],
        'vidres': vidres or [],
        'passpartout': passpartout or [],
        'encolat_pro': encolat_pro or [],
    }

    def fake_query(sql, args=(), one=False):
        for taula, rows in data.items():
            if f'FROM {taula}' in sql:
                return rows
        return []

    monkeypatch.setattr(app, 'query', fake_query)
    return app._auditoria_tarifes()


def _has(findings, sev=None, ref=None, camp=None):
    return any(
        (sev is None or f['sev'] == sev)
        and (ref is None or f['ref'] == ref)
        and (camp is None or f['camp'] == camp)
        for f in findings
    )


# ── helpers ───────────────────────────────────────────────────────────

def test_median():
    assert app._au_median([1, 2, 3]) == 2
    assert app._au_median([1, 2, 3, 4]) == 2.5
    assert app._au_median([]) is None


# ── moldures ──────────────────────────────────────────────────────────

def test_moldura_sense_preu_es_critica(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'M1', 'preu_taller': 0, 'preu_cost': None, 'gruix': 2}])
    assert _has(f, sev='critical', ref='M1', camp='preu')


def test_moldura_descatalogada_sense_preu_no_avisa(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'M1', 'preu_taller': 0, 'preu_cost': None, 'gruix': 2, 'descatalogada': True}])
    assert not _has(f, ref='M1')


def test_moldura_cost_per_sobre_preu(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'M2', 'preu_taller': 5.0, 'preu_cost': 9.0, 'gruix': 2}])
    assert _has(f, sev='warning', ref='M2', camp='preu_cost')


def test_moldura_sense_preu_cost_v2(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'M3', 'preu_taller': 5.0, 'preu_cost': None, 'gruix': 2}])
    assert _has(f, sev='warning', ref='M3', camp='preu_cost')


def test_moldura_gruix_buit(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'M4', 'preu_taller': 5.0, 'preu_cost': 3.0, 'gruix': 0}])
    assert _has(f, sev='warning', ref='M4', camp='gruix')


def test_moldura_merma_atipica(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'M5', 'preu_taller': 5.0, 'preu_cost': 3.0, 'gruix': 2, 'merma_pct': 80}])
    assert _has(f, sev='warning', ref='M5', camp='merma_pct')


def test_moldura_ref2_duplicada(monkeypatch):
    f, r = _run(monkeypatch, moldures=[
        {'referencia': 'A', 'preu_taller': 5, 'preu_cost': 3, 'gruix': 2, 'ref2': 'X99'},
        {'referencia': 'B', 'preu_taller': 5, 'preu_cost': 3, 'gruix': 2, 'ref2': 'X99'},
    ])
    assert _has(f, sev='info', camp='ref2')


def test_moldura_correcta_no_dispara(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'OK', 'preu_taller': 6.0, 'preu_cost': 3.0, 'gruix': 2.0, 'merma_pct': 10, 'minim_cm': 100}])
    assert not _has(f, ref='OK')


# ── vidres / passpartout / encolat ────────────────────────────────────

def test_vidre_cost_per_sobre_preu(monkeypatch):
    f, r = _run(monkeypatch, vidres=[{'referencia': 'V1', 'preu': 4.0, 'preu_cost': 7.0}])
    assert _has(f, sev='warning', ref='V1', camp='preu_cost')


def test_encolat_sense_preu_critica(monkeypatch):
    f, r = _run(monkeypatch, encolat_pro=[{'referencia': 'E1', 'preu': None, 'preu_cost': 0}])
    assert _has(f, sev='critical', ref='E1')


# ── outliers ──────────────────────────────────────────────────────────

def test_outlier_alt_detectat(monkeypatch):
    # 12 motllures normals (~3) + 1 disparada (300)
    mols = [{'referencia': f'N{i}', 'preu_taller': 6, 'preu_cost': 3.0, 'gruix': 2} for i in range(12)]
    mols.append({'referencia': 'BIG', 'preu_taller': 600, 'preu_cost': 300.0, 'gruix': 2})
    f, r = _run(monkeypatch, moldures=mols)
    assert _has(f, ref='BIG', camp='preu_cost')


def test_resum_compta_severitats(monkeypatch):
    f, r = _run(monkeypatch, moldures=[{'referencia': 'M1', 'preu_taller': 0, 'preu_cost': None, 'gruix': 2}])
    assert r['critical'] >= 1
    assert r['total'] == len(f)

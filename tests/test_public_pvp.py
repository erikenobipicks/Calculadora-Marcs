"""Tests de l'API pública de preus PVP (/api/public/pvp/*) i del motor de preu
compartit.

Cobreix:
- el token dedicat X-Pvp-Token (constant + rotació + independència del bridge),
- la conversió PVD → PVP + IVA (_pvp_from_pvd) i el marge per defecte,
- l'endpoint: rebuig 403, errors 400/404 i que la resposta MAI filtra cost/PVD.
"""
import pytest

import app


# ── _pvp_token_ok: token dedicat ───────────────────────────────────────────
def test_pvp_token_accepts_primary(monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    monkeypatch.delenv('PUBLIC_PVP_TOKEN_NEXT', raising=False)
    assert app._pvp_token_ok('pvp-secret') is True


def test_pvp_token_rejects_wrong(monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    monkeypatch.delenv('PUBLIC_PVP_TOKEN_NEXT', raising=False)
    assert app._pvp_token_ok('nope') is False


def test_pvp_token_empty_and_no_env(monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    assert app._pvp_token_ok('') is False
    assert app._pvp_token_ok(None) is False
    monkeypatch.delenv('PUBLIC_PVP_TOKEN', raising=False)
    assert app._pvp_token_ok('pvp-secret') is False  # fail-closed sense env


def test_pvp_token_rotation(monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'old')
    monkeypatch.setenv('PUBLIC_PVP_TOKEN_NEXT', 'new')
    assert app._pvp_token_ok('old') is True
    assert app._pvp_token_ok('new') is True
    assert app._pvp_token_ok('x') is False


def test_pvp_token_independent_from_bridge(monkeypatch):
    """El token PVP i el de bridge són independents: cap serveix per a l'altre."""
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'bridge-secret')
    monkeypatch.delenv('PUBLIC_PVP_TOKEN_NEXT', raising=False)
    monkeypatch.delenv('PUBLIC_BRIDGE_TOKEN_NEXT', raising=False)
    assert app._pvp_token_ok('bridge-secret') is False
    assert app._bridge_token_ok('pvp-secret') is False


# ── _pvp_from_pvd: PVD → PVP + IVA ─────────────────────────────────────────
def test_pvp_from_pvd_basic():
    assert app._pvp_from_pvd(20.0, 60.0) == {'pvp_net': 32.0, 'iva': 6.72, 'pvp_total': 38.72}


def test_pvp_from_pvd_zero_margin():
    assert app._pvp_from_pvd(100.0, 0.0) == {'pvp_net': 100.0, 'iva': 21.0, 'pvp_total': 121.0}


def test_pvp_from_pvd_rounding():
    r = app._pvp_from_pvd(9.99, 33.0)  # 9.99 * 1.33 = 13.2867 -> 13.29
    assert r['pvp_net'] == 13.29
    assert r['iva'] == round(13.29 * 0.21, 2)
    assert r['pvp_total'] == round(13.29 + r['iva'], 2)


# ── _marge_public_pct ──────────────────────────────────────────────────────
def test_marge_public_default(monkeypatch):
    monkeypatch.setattr(app, 'get_config_value', lambda clau, default=None: default)
    assert app._marge_public_pct() == 60.0


def test_marge_public_from_config(monkeypatch):
    monkeypatch.setattr(app, 'get_config_value', lambda clau, default=None: '45')
    assert app._marge_public_pct() == 45.0


# ── Endpoint /api/public/pvp/compute ───────────────────────────────────────
@pytest.fixture
def client():
    app.app.config['TESTING'] = True
    return app.app.test_client()


def test_pvp_endpoint_rejects_missing_token(client, monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    r = client.get('/api/public/pvp/compute?kind=frame&width_cm=30&height_cm=40&moldura_id=X')
    assert r.status_code == 403


def test_pvp_endpoint_rejects_wrong_token(client, monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    r = client.get('/api/public/pvp/compute?kind=frame&width_cm=30&height_cm=40&moldura_id=X',
                   headers={'X-Pvp-Token': 'wrong'})
    assert r.status_code == 403


def test_pvp_endpoint_bridge_token_not_accepted(client, monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'bridge-secret')
    r = client.get('/api/public/pvp/compute?kind=frame&width_cm=30&height_cm=40&moldura_id=X',
                   headers={'X-Pvp-Token': 'bridge-secret'})
    assert r.status_code == 403


def test_pvp_endpoint_unknown_kind(client, monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    r = client.get('/api/public/pvp/compute?kind=banana&width_cm=30&height_cm=40',
                   headers={'X-Pvp-Token': 'pvp-secret'})
    assert r.status_code == 400
    assert r.get_json()['error'] == 'unknown_kind'


def test_pvp_endpoint_invalid_size(client, monkeypatch):
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    r = client.get('/api/public/pvp/compute?kind=frame&width_cm=0&height_cm=40&moldura_id=X',
                   headers={'X-Pvp-Token': 'pvp-secret'})
    assert r.status_code == 400
    assert r.get_json()['error'] == 'invalid_size'


def test_pvp_endpoint_math_and_no_cost_leak(client, monkeypatch):
    """base PVD 10 × qty 2 = 20; marge 60% → net 32,00 · IVA 6,72 · total 38,72.
    I la resposta no ha de filtrar mai cost/PVD/breakdown."""
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    monkeypatch.setattr(app, '_public_base_price_unit',
                        lambda *a, **kw: (10.0, {'moldura': 10.0, 'cost': 4.0}, None))
    r = client.get('/api/public/pvp/compute?kind=frame&width_cm=30&height_cm=40&moldura_id=X&qty=2',
                   headers={'X-Pvp-Token': 'pvp-secret'})
    assert r.status_code == 200
    d = r.get_json()
    assert d['ok'] is True
    assert d['pvp_net'] == 32.0
    assert d['iva'] == 6.72
    assert d['pvp_total'] == 38.72
    assert d['marge_pct'] == 60.0
    assert d['qty'] == 2
    for leaked in ('cost', 'pvd', 'base_price', 'breakdown'):
        assert leaked not in d
    body = r.get_data(as_text=True)
    assert '"cost"' not in body and '"breakdown"' not in body and '"base_price"' not in body


def test_pvp_endpoint_moldura_not_found_is_404(client, monkeypatch):
    """Amb query()->[] (conftest), un marc inexistent dona 404, no 403."""
    monkeypatch.setenv('PUBLIC_PVP_TOKEN', 'pvp-secret')
    r = client.get('/api/public/pvp/compute?kind=frame&width_cm=30&height_cm=40&moldura_id=NOEXIST',
                   headers={'X-Pvp-Token': 'pvp-secret'})
    assert r.status_code == 404
    assert r.get_json()['error'] == 'moldura_not_found'

"""Tests de l'enduriment del bridge token (FASE 1 / C-1).

Cobreixen la verificació en temps constant amb suport de rotació
(PUBLIC_BRIDGE_TOKEN + PUBLIC_BRIDGE_TOKEN_NEXT) i l'allowlist d'IPs
opcional (BRIDGE_ALLOWED_IPS), a més del rebuig 403 a nivell d'endpoint.
"""
import importlib

import pytest

import app


# ── _bridge_token_ok: comparació constant + rotació ────────────────────────
def test_token_ok_accepts_primary(monkeypatch):
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'primary-secret')
    monkeypatch.delenv('PUBLIC_BRIDGE_TOKEN_NEXT', raising=False)
    assert app._bridge_token_ok('primary-secret') is True


def test_token_ok_rejects_wrong(monkeypatch):
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'primary-secret')
    monkeypatch.delenv('PUBLIC_BRIDGE_TOKEN_NEXT', raising=False)
    assert app._bridge_token_ok('nope') is False


def test_token_ok_empty_is_false(monkeypatch):
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'primary-secret')
    assert app._bridge_token_ok('') is False
    assert app._bridge_token_ok(None) is False


def test_token_ok_no_env_is_false(monkeypatch):
    """Sense cap token configurat, mai s'autoritza (fail-closed)."""
    monkeypatch.delenv('PUBLIC_BRIDGE_TOKEN', raising=False)
    monkeypatch.delenv('PUBLIC_BRIDGE_TOKEN_NEXT', raising=False)
    assert app._bridge_token_ok('whatever') is False


def test_token_ok_rotation_accepts_both(monkeypatch):
    """Durant una rotació, el token vell i el nou són vàlids alhora."""
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'old-secret')
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN_NEXT', 'new-secret')
    assert app._bridge_token_ok('old-secret') is True
    assert app._bridge_token_ok('new-secret') is True
    assert app._bridge_token_ok('other') is False


# ── _bridge_ip_allowed: allowlist opcional ─────────────────────────────────
def test_ip_allowed_no_env_is_open(monkeypatch):
    """Sense BRIDGE_ALLOWED_IPS no es restringeix res (retrocompatible)."""
    monkeypatch.delenv('BRIDGE_ALLOWED_IPS', raising=False)
    with app.app.test_request_context(environ_overrides={'REMOTE_ADDR': '9.9.9.9'}):
        assert app._bridge_ip_allowed() is True


def test_ip_allowed_match(monkeypatch):
    monkeypatch.setenv('BRIDGE_ALLOWED_IPS', '1.2.3.4, 10.0.0.0/8')
    with app.app.test_request_context(environ_overrides={'REMOTE_ADDR': '10.0.0.5'}):
        assert app._bridge_ip_allowed() is True
    with app.app.test_request_context(environ_overrides={'REMOTE_ADDR': '1.2.3.4'}):
        assert app._bridge_ip_allowed() is True


def test_ip_allowed_no_match(monkeypatch):
    monkeypatch.setenv('BRIDGE_ALLOWED_IPS', '1.2.3.4, 10.0.0.0/8')
    with app.app.test_request_context(environ_overrides={'REMOTE_ADDR': '8.8.8.8'}):
        assert app._bridge_ip_allowed() is False


def test_ip_allowed_bad_remote_is_false(monkeypatch):
    monkeypatch.setenv('BRIDGE_ALLOWED_IPS', '1.2.3.4')
    with app.app.test_request_context(environ_overrides={'REMOTE_ADDR': ''}):
        assert app._bridge_ip_allowed() is False


# ── Rebuig a nivell d'endpoint ─────────────────────────────────────────────
@pytest.fixture
def client():
    app.app.config['TESTING'] = True
    return app.app.test_client()


def test_endpoint_rejects_missing_token(client, monkeypatch):
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'primary-secret')
    resp = client.get('/api/public/clients-habituals')
    assert resp.status_code == 403


def test_endpoint_rejects_wrong_token(client, monkeypatch):
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'primary-secret')
    resp = client.get('/api/public/clients-habituals',
                      headers={'X-Bridge-Token': 'wrong'})
    assert resp.status_code == 403


def test_endpoint_accepts_valid_token(client, monkeypatch):
    """Amb el token correcte ja no és 403 (no comprovem el cos, que depèn de
    la BD; només que la barrera d'autenticació deixa passar)."""
    monkeypatch.setenv('PUBLIC_BRIDGE_TOKEN', 'primary-secret')
    resp = client.get('/api/public/clients-habituals',
                      headers={'X-Bridge-Token': 'primary-secret'})
    assert resp.status_code != 403

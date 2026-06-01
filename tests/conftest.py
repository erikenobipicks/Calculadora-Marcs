import os
import sys

# app.py crida sys.exit si no hi ha SECRET_KEY a l'import. Posem un valor
# fals abans que cap test importi `app`.
os.environ.setdefault('SECRET_KEY', 'test-only-not-for-production')

# Permet `import app` quan es corre `pytest tests/` des de l'arrel.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

import pytest


@pytest.fixture(autouse=True)
def _no_db_no_config(monkeypatch):
    """Força a totes les funcions calcular_cost_* a anar pel camí fórmula:
    - `query` retorna sempre [] → cap fila a la taula → cap match.
    - `get_config_value` retorna sempre el default → defaults seedats a init_db().

    Si un test necessita simular files concretes a una taula, pot
    sobreescriure el mock de `query` amb el seu propi monkeypatch.
    """
    import app
    monkeypatch.setattr(app, 'query', lambda *a, **kw: [])
    monkeypatch.setattr(app, 'get_config_value', lambda clau, default=None: default)

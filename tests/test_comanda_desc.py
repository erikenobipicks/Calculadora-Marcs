"""La descripció d'una comanda usa el concepte complet desat quan hi és, així el
document mostra tots els conceptes (encolat foam, vidre…); si no, el reconstrueix."""
import app


def test_desc_prefers_concepte():
    com = {'concepte': 'Marc 30x40 + Encolat + Vidre', 'marc_principal': 'REF1',
           'encolat': '', 'vidre': '', 'tipus_peca': ''}
    assert app._comanda_linia_desc(com) == 'Marc 30x40 + Encolat + Vidre'


def test_desc_fallback_when_no_concepte():
    com = {'concepte': '', 'marc_principal': 'M25', 'pre_marc': '', 'passpartout': '',
           'vidre': 'V1', 'encolat': 'FOAM', 'impressio': '', 'revers_peu': '',
           'tipus_peca': ''}
    d = app._comanda_linia_desc(com)
    assert 'M25' in d and 'V1' in d and 'FOAM' in d

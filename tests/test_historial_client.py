"""El nom del client a l'historial mostra també el nom comercial."""
import app


def test_hist_client_display_anteposa_comercial():
    assert app._hist_client_display('Fotografia Poblet', 'Rosa') == 'Fotografia Poblet · Rosa'


def test_hist_client_display_no_duplica():
    # Si el comercial ja és dins del nom desat (pressupostos nous), no el duplica.
    assert app._hist_client_display('Poblet', 'Poblet (Rosa)') == 'Poblet (Rosa)'


def test_hist_client_display_fallbacks():
    assert app._hist_client_display('', 'Rosa') == 'Rosa'
    assert app._hist_client_display('X', '') == 'X'
    assert app._hist_client_display(None, None) == ''

"""Recàrrec d'equivalència a Factura Directa.

A FacturaDirecta el recàrrec d'equivalència es representa com un impost extra a
la línia, JUNT amb l'IVA. El codi real del compte és 'S_IVA_RE_5.2' (verificat
en documents existents; FD el rebutjava abans perquè fèiem servir 'S_REQ_52',
inexistent). FD no l'aplica sol als documents fets per API, així que
_fd_line_tax l'afegeix per als clients en règim de recàrrec.
"""
import app


def test_fd_line_tax_nomes_iva():
    assert app._fd_line_tax(False) == [app._FD_IVA_CODE]


def test_fd_line_tax_afegeix_recarrec():
    assert app._fd_line_tax(True) == [app._FD_IVA_CODE, app._FD_RE_CODE]
    assert app._FD_RE_CODE in app._fd_line_tax(True)


def test_fd_re_code_default():
    # Codi verificat en documents reals del compte de FacturaDirecta.
    assert app._FD_RE_CODE == 'S_IVA_RE_5.2'

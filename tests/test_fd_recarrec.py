"""Correcció del recàrrec d'equivalència a Factura Directa.

FacturaDirecta rebutja el recàrrec com a impost de línia amb l'error
"Impuesto desconocido: 'S_REQ_52'": el recàrrec d'equivalència és un RÈGIM del
contacte, no una taxa per línia. Per tant _fd_line_tax mai hi afegeix el codi de
recàrrec (només l'IVA); FD aplica el recàrrec des del règim del contacte.
"""
import app


def test_fd_line_tax_nomes_iva():
    assert app._fd_line_tax(False) == [app._FD_IVA_CODE]


def test_fd_line_tax_recarrec_no_afegeix_codi():
    # Encara que el client sigui de recàrrec, la línia porta només l'IVA.
    assert app._fd_line_tax(True) == [app._FD_IVA_CODE]
    assert app._FD_RE_CODE not in app._fd_line_tax(True)

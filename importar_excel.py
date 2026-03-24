"""
Executa aquest script UNA VEGADA per importar el catàleg de l'Excel a la base de dades.
Col·loca l'Excel al mateix directori i executa:
    python3 importar_excel.py nom_del_fitxer.xlsx
"""
import sys, os, sqlite3
import openpyxl

DB = os.path.join(os.path.dirname(__file__), 'objectiu.db')

def importar(path_excel):
    wb = openpyxl.load_workbook(path_excel, data_only=True)
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    # Moldures
    ws = wb['Lista Molduras']
    count_m = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        ref = row[0]
        if not ref or str(ref).strip() in ('', 'Referencia', '-'):
            continue
        ref = str(ref).strip()
        preu   = float(row[1]) if row[1] else 0
        gruix  = float(row[2]) if row[2] else 0
        cost   = float(row[3]) if row[3] else 0
        prov   = str(row[4]).strip() if row[4] else ''
        ref2   = str(row[5]).strip() if row[5] else ''
        ubi    = str(row[6]).strip() if row[6] else ''
        desc   = str(row[7]).strip() if row[7] else ''
        c.execute('INSERT OR REPLACE INTO moldures VALUES (?,?,?,?,?,?,?,?,NULL)',
                  [ref, preu, gruix, cost, prov, ref2, ubi, desc])
        count_m += 1

    # Vidres
    ws = wb['Vidres']
    count_v = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        ref, preu = row[0], row[1]
        if not ref or not preu: continue
        c.execute('INSERT OR REPLACE INTO vidres VALUES (?,?)', [str(ref).strip(), float(preu)])
        count_v += 1

    # Passpartout
    ws = wb['Passpertout']
    count_p = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        ref, preu = row[0], row[1]
        if not ref or not preu: continue
        c.execute('INSERT OR REPLACE INTO passpartout VALUES (?,?)', [str(ref).strip(), float(preu)])
        count_p += 1

    # Encolat/Pro
    ws = wb['Encolado_Pro']
    count_e = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        ref, preu = row[0], row[1]
        if not ref or not preu: continue
        c.execute('INSERT OR REPLACE INTO encolat_pro VALUES (?,?)', [str(ref).strip(), float(preu)])
        count_e += 1

    conn.commit()
    conn.close()
    print(f"Import complet:")
    print(f"  Moldures:    {count_m}")
    print(f"  Vidres:      {count_v}")
    print(f"  Passpartout: {count_p}")
    print(f"  Encolat/Pro: {count_e}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Ús: python3 importar_excel.py fitxer.xlsx")
        sys.exit(1)
    importar(sys.argv[1])

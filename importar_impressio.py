"""
Importa les tarifes d'impressió fotogràfica a la base de dades.
Executa: python importar_impressio.py
"""
import os

DATABASE_URL = os.environ.get('DATABASE_URL', '')
if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
USE_PG = bool(DATABASE_URL)

if USE_PG:
    import psycopg2
    conn = psycopg2.connect(DATABASE_URL)
    conn.autocommit = False
    cur = conn.cursor()
    def upsert(ref, desc, preu):
        cur.execute("INSERT INTO impressio (referencia,descripcio,preu) VALUES (%s,%s,%s) ON CONFLICT (referencia) DO UPDATE SET descripcio=EXCLUDED.descripcio, preu=EXCLUDED.preu", [ref, desc, preu])
else:
    import sqlite3
    DB = os.path.join(os.path.dirname(__file__), 'objectiu.db')
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    def upsert(ref, desc, preu):
        cur.execute("INSERT OR REPLACE INTO impressio VALUES (?,?,?)", [ref, desc, preu])

# Tarifes d'impressió fotogràfica
tarifes = [
    ("IMP7x10",    "7x10 cm",    0.61),
    ("IMP10x15",   "10x15 cm",   0.61),
    ("IMP13x13",   "13x13 cm",   0.79),
    ("IMP13x18",   "13x18 cm",   0.87),
    ("IMP15x15",   "15x15 cm",   0.87),
    ("IMP15x20",   "15x20 cm",   0.87),
    ("IMP15x23",   "15x23 cm",   1.05),
    ("IMP20x20",   "20x20 cm",   1.35),
    ("IMP20x25",   "20x25 cm",   1.35),
    ("IMP20x30",   "20x30 cm",   1.57),
    ("IMP25x25",   "25x25 cm",   1.57),
    ("IMP25x30",   "25x30 cm",   1.87),
    ("IMP25x38",   "25x38 cm",   2.35),
    ("IMP30x30",   "30x30 cm",   2.35),
    ("IMP30x40",   "30x40 cm",   3.60),
    ("IMP30x45",   "30x45 cm",   3.60),
    ("IMP40x40",   "40x40 cm",   4.75),
    ("IMP40x50",   "40x50 cm",   5.10),
    ("IMP40x60",   "40x60 cm",   5.70),
    ("IMP50x50",   "50x50 cm",   7.15),
    ("IMP50x60",   "50x60 cm",   7.90),
    ("IMP50x70",   "50x70 cm",   8.75),
    ("IMP50x75",   "50x75 cm",   9.35),
    ("IMP60x60",   "60x60 cm",   9.35),
    ("IMP60x80",   "60x80 cm",  11.90),
    ("IMP60x90",   "60x90 cm",  13.20),
    ("IMP70x70",   "70x70 cm",  13.20),
    ("IMP70x90",   "70x90 cm",  16.40),
    ("IMP70x100",  "70x100 cm", 18.50),
    ("IMP80x80",   "80x80 cm",  18.50),
    ("IMP80x100",  "80x100 cm", 22.00),
    ("IMP80x120",  "80x120 cm", 26.50),
    ("IMP90x90",   "90x90 cm",  23.00),
    ("IMP90x120",  "90x120 cm", 29.50),
    ("IMP100x100", "100x100 cm",32.00),
    ("IMP100x120", "100x120 cm",36.00),
    ("IMP100x130", "100x130 cm",39.00),
    ("IMP100x140", "100x140 cm",42.00),
    ("IMP100x150", "100x150 cm",45.00),
    ("IMP110x140", "110x140 cm",51.00),
    ("IMP110x150", "110x150 cm",55.00),
    ("IMP110x160", "110x160 cm",66.00),
    ("IMP120x150", "120x150 cm",60.00),
    ("IMP120x160", "120x160 cm",70.00),
    ("IMP120x180", "120x180 cm",80.00),
]

for ref, desc, preu in tarifes:
    upsert(ref, desc, preu)

conn.commit()
conn.close()
print(f"Importades {len(tarifes)} tarifes d'impressió OK")

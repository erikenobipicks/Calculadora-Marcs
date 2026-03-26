"""
Importa les tarifes d'impressió fotogràfica a la base de dades.
Executa: python importar_impressio.py
"""
import os

DATABASE_URL = os.environ.get('DATABASE_URL', '')
if DATABASE_URL.startswith('postgres://'): DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
USE_PG = bool(DATABASE_URL)

if USE_PG:
    import psycopg2
    conn = psycopg2.connect(DATABASE_URL); conn.autocommit=False; cur=conn.cursor()
    def upsert(ref,desc,preu): cur.execute("INSERT INTO impressio (referencia,descripcio,preu) VALUES (%s,%s,%s) ON CONFLICT (referencia) DO UPDATE SET descripcio=EXCLUDED.descripcio, preu=EXCLUDED.preu",[ref,desc,preu])
else:
    import sqlite3
    DB=os.path.join(os.path.dirname(__file__),'objectiu.db'); conn=sqlite3.connect(DB); cur=conn.cursor()
    def upsert(ref,desc,preu): cur.execute("INSERT OR REPLACE INTO impressio VALUES (?,?,?)",[ref,desc,preu])

tarifes = [
    ("IMP7x10","7x10 cm",0.61),("IMP10x15","10x15 cm",0.61),
    ("IMP13x13","13x13 cm",0.79),("IMP13x18","13x18 cm",0.87),
    ("IMP15x15","15x15 cm",0.87),("IMP15x20","15x20 cm",0.87),
    ("IMP18x24","18x24 cm",1.40),("IMP20x20","20x20 cm",1.54),
    ("IMP20x25","20x25 cm",1.57),("IMP20x30","20x30 cm",1.94),
    ("IMP21x30","21x30 cm",2.10),("IMP24x30","24x30 cm",2.42),
    ("IMP24x36","24x36 cm",2.80),("IMP28x35","28x35 cm",3.00),
    ("IMP30x30","30x30 cm",3.30),("IMP30x40","30x40 cm",3.60),
    ("IMP30x45","30x45 cm",3.80),("IMP30x50","30x50 cm",4.90),
    ("IMP35x50","35x50 cm",5.80),("IMP30x60","30x60 cm",5.90),
    ("IMP40x50","40x50 cm",6.40),("IMP40x60","40x60 cm",7.80),
    ("IMP50x50","50x50 cm",8.50),("IMP50x60","50x60 cm",9.20),
    ("IMP50x70","50x70 cm",10.50),("IMP50x80","50x80 cm",13.50),
    ("IMP60x100","60x100 cm",24.00),("IMP70x100","70x100 cm",26.00),
    ("IMP80x100","80x100 cm",28.00),("IMP80x110","80x110 cm",30.00),
    ("IMP80x120","80x120 cm",32.00),("IMP80x180","80x180 cm",45.00),
    ("IMP80x200","80x200 cm",52.00),("IMP90x100","90x100 cm",34.00),
    ("IMP90x110","90x110 cm",35.00),("IMP90x150","90x150 cm",39.00),
    ("IMP90x180","90x180 cm",45.00),("IMP90x200","90x200 cm",52.00),
    ("IMP100x100","100x100 cm",32.00),("IMP100x150","100x150 cm",49.00),
    ("IMP100x180","100x180 cm",55.00),("IMP100x200","100x200 cm",70.00),
    ("IMP110x110","110x110 cm",42.00),("IMP110x130","110x130 cm",52.00),
    ("IMP110x160","110x160 cm",66.00),
]

# Delete old rates first to avoid stale data
if USE_PG: cur.execute("DELETE FROM impressio")
else: cur.execute("DELETE FROM impressio")

for ref,desc,preu in tarifes: upsert(ref,desc,preu)
conn.commit(); conn.close()
print(f"Importades {len(tarifes)} tarifes OK")
print(f"50x60 = {[t for t in tarifes if '50x60' in t[0]]}")

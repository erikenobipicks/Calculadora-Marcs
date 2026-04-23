import os
import sqlite3
import psycopg2
import psycopg2.extras

# ── Configuració ─────────────────────────────────────────────
SQLITE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "objectiu.db"
)

PG_URL = os.environ.get("DATABASE_URL")
if not PG_URL:
    raise SystemExit(
        "DATABASE_URL no definida. "
        "Executa: railway run python migrate_sqlite_to_pg.py"
    )
if PG_URL.startswith("postgres://"):
    PG_URL = PG_URL.replace("postgres://", "postgresql://", 1)
if "sslmode" not in PG_URL:
    PG_URL += ("&" if "?" in PG_URL else "?") + "sslmode=require"

# ── Ordre respectant FK ───────────────────────────────────────
TAULES = [
    "config",
    "usuaris",
    "moldures",
    "vidres",
    "encolat_pro",
    "passpartout",
    "proeco",
    "impressio",
    "comandes",
    "feedback",
    "historial_preus_cost",
]

# ── Taules amb SERIAL (seqüències a actualitzar) ─────────────
SEQS = {
    "usuaris":              "usuaris_id_seq",
    "comandes":             "comandes_id_seq",
    "feedback":             "feedback_id_seq",
    "historial_preus_cost": "historial_preus_cost_id_seq",
}


def migrar():
    sqlite = sqlite3.connect(SQLITE_PATH)
    sqlite.row_factory = sqlite3.Row
    pg = psycopg2.connect(PG_URL,
                          cursor_factory=psycopg2.extras.RealDictCursor)

    total_ok  = 0
    total_err = 0

    for taula in TAULES:
        # Comprovar que la taula existeix al SQLite
        exists = sqlite.execute("""
            SELECT COUNT(*) FROM sqlite_master
            WHERE type='table' AND name=?
        """, [taula]).fetchone()[0]

        if not exists:
            print(f"{taula}: no existeix al SQLite, saltant")
            continue

        files = sqlite.execute(f"SELECT * FROM {taula}").fetchall()
        if not files:
            print(f"{taula}: 0 files, saltant")
            continue

        cols   = list(files[0].keys())
        colstr = ", ".join(cols)
        marks  = ", ".join(["%s"] * len(cols))

        ok = err = 0
        with pg.cursor() as cur:
            for fila in files:
                try:
                    cur.execute(
                        f"INSERT INTO {taula} ({colstr}) "
                        f"VALUES ({marks}) "
                        f"ON CONFLICT DO NOTHING",
                        tuple(fila)
                    )
                    ok += 1
                except Exception as e:
                    pg.rollback()
                    print(f"  ⚠️  {taula} fila {dict(fila).get('id','?')}: {e}")
                    err += 1
                    continue
        pg.commit()
        print(f"{taula}: {ok} ok, {err} errors")
        total_ok  += ok
        total_err += err

    # ── Actualitzar seqüències SERIAL ────────────────────────
    print("\nActualitzant seqüències SERIAL...")
    with pg.cursor() as cur:
        for taula, seq in SEQS.items():
            try:
                cur.execute(f"""
                    SELECT setval(
                        '{seq}',
                        COALESCE((SELECT MAX(id) FROM {taula}), 1)
                    )
                """)
                result = cur.fetchone()
                print(f"  {seq} → {list(result.values())[0]}")
            except Exception as e:
                print(f"  ⚠️  {seq}: {e}")
    pg.commit()

    print(f"\nMigració completada: {total_ok} files OK, {total_err} errors")
    sqlite.close()
    pg.close()


if __name__ == "__main__":
    migrar()

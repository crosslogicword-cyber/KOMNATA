import os
import sqlite3
from contextlib import closing
import psycopg2

SQLITE_PATH = "komnata.db"
DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    raise SystemExit("BRAK DATABASE_URL. Najpierw: set -a && source .env && set +a")

TABLES_ORDER = [
    "layouts",
    "sectors",
    "racks",
    "rack_slots",
    "items",
    "item_history",
]

UPSERTS = {
    "layouts": """
        INSERT INTO layouts (id, name, start_number, end_number, letters, is_active, description)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            name = EXCLUDED.name,
            start_number = EXCLUDED.start_number,
            end_number = EXCLUDED.end_number,
            letters = EXCLUDED.letters,
            is_active = EXCLUDED.is_active,
            description = EXCLUDED.description
    """,
    "sectors": """
        INSERT INTO sectors (id, name, is_active, created_at)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            name = EXCLUDED.name,
            is_active = EXCLUDED.is_active,
            created_at = EXCLUDED.created_at
    """,
    "racks": """
        INSERT INTO racks (id, name, is_active, created_at, layout_id)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            name = EXCLUDED.name,
            is_active = EXCLUDED.is_active,
            created_at = EXCLUDED.created_at,
            layout_id = EXCLUDED.layout_id
    """,
    "rack_slots": """
        INSERT INTO rack_slots (id, rack_id, code, is_active, sort_order)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            rack_id = EXCLUDED.rack_id,
            code = EXCLUDED.code,
            is_active = EXCLUDED.is_active,
            sort_order = EXCLUDED.sort_order
    """,
    "items": """
        INSERT INTO items (id, product_name, quantity, notes, sector_id, rack_id, created_at, slot_id, extra_field, updated_at)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            product_name = EXCLUDED.product_name,
            quantity = EXCLUDED.quantity,
            notes = EXCLUDED.notes,
            sector_id = EXCLUDED.sector_id,
            rack_id = EXCLUDED.rack_id,
            created_at = EXCLUDED.created_at,
            slot_id = EXCLUDED.slot_id,
            extra_field = EXCLUDED.extra_field,
            updated_at = EXCLUDED.updated_at
    """,
    "item_history": """
        INSERT INTO item_history (id, item_id, action, details, changed_at)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (id) DO UPDATE SET
            item_id = EXCLUDED.item_id,
            action = EXCLUDED.action,
            details = EXCLUDED.details,
            changed_at = EXCLUDED.changed_at
    """,
}

def sqlite_rows(conn, table):
    cur = conn.execute(f"SELECT * FROM {table} ORDER BY id")
    return cur.fetchall()

def boolify(v):
    if v in (None, "", b""):
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    s = str(v).strip().lower()
    if s in {"1", "true", "t", "yes"}:
        return True
    if s in {"0", "false", "f", "no"}:
        return False
    return v

def normalize_row(table, row):
    r = dict(row)
    if "is_active" in r:
        r["is_active"] = boolify(r["is_active"])
    return r

def tuple_for_table(table, r):
    if table == "layouts":
        return (r.get("id"), r.get("name"), r.get("start_number"), r.get("end_number"), r.get("letters"), r.get("is_active"), r.get("description"))
    if table == "sectors":
        return (r.get("id"), r.get("name"), r.get("is_active"), r.get("created_at"))
    if table == "racks":
        return (r.get("id"), r.get("name"), r.get("is_active"), r.get("created_at"), r.get("layout_id"))
    if table == "rack_slots":
        return (r.get("id"), r.get("rack_id"), r.get("code"), r.get("is_active"), r.get("sort_order"))
    if table == "items":
        return (r.get("id"), r.get("product_name"), r.get("quantity"), r.get("notes"), r.get("sector_id"), r.get("rack_id"), r.get("created_at"), r.get("slot_id"), r.get("extra_field"), r.get("updated_at"))
    if table == "item_history":
        return (r.get("id"), r.get("item_id"), r.get("action"), r.get("details"), r.get("changed_at"))
    raise ValueError(table)

def set_sequence(pg_conn, table):
    with pg_conn.cursor() as cur:
        cur.execute("SELECT pg_get_serial_sequence(%s, 'id')", (table,))
        row = cur.fetchone()
        if not row or not row[0]:
            return
        seq = row[0]
        cur.execute(f"SELECT COALESCE(MAX(id), 1) FROM {table}")
        max_id = cur.fetchone()[0] or 1
        cur.execute("SELECT setval(%s, %s, true)", (seq, max_id))

def main():
    sq = sqlite3.connect(SQLITE_PATH)
    sq.row_factory = sqlite3.Row
    pg = psycopg2.connect(DATABASE_URL, sslmode="require")
    pg.autocommit = False

    try:
        with closing(sq), closing(pg):
            for table in TABLES_ORDER:
                rows = sqlite_rows(sq, table)
                print(f"{table}: SQLite rows = {len(rows)}")
                done = 0
                for row in rows:
                    r = normalize_row(table, row)
                    values = tuple_for_table(table, r)
                    with pg.cursor() as cur:
                        cur.execute(UPSERTS[table], values)
                    done += 1
                pg.commit()
                set_sequence(pg, table)
                pg.commit()
                print(f"{table}: przeniesiono = {done}")
            print("OK: migracja zakończona")
    except Exception:
        pg.rollback()
        raise

if __name__ == "__main__":
    main()

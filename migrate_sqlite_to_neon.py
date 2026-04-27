import os
import sqlite3
from dotenv import load_dotenv
import psycopg2

load_dotenv(".env")

sqlite_conn = sqlite3.connect("komnata.db")
sqlite_conn.row_factory = sqlite3.Row

pg_conn = psycopg2.connect(os.getenv("DATABASE_URL"))
pg_conn.autocommit = False

TABLES = [
    ("sectors", ["id", "name", "is_active", "created_at"]),
    ("layouts", ["id", "name", "start_number", "end_number", "letters", "description", "is_active", "created_at"]),
    ("racks", ["id", "name", "is_active", "created_at", "layout_id"]),
    ("rack_slots", ["id", "rack_id", "code", "is_active", "created_at", "sort_order"]),
    ("items", ["id", "product_name", "quantity", "notes", "sector_id", "rack_id", "created_at", "slot_id", "extra_field", "updated_at"]),
    ("item_history", ["id", "item_id", "action", "details", "changed_at"]),
]

def copy_table(table, columns):
    s_cur = sqlite_conn.cursor()
    p_cur = pg_conn.cursor()

    s_cur.execute(f"SELECT {', '.join(columns)} FROM {table}")
    rows = s_cur.fetchall()

    if not rows:
        print(f"{table}: 0")
        p_cur.close()
        s_cur.close()
        return

    placeholders = ", ".join(["%s"] * len(columns))
    col_list = ", ".join(columns)

    for row in rows:
        values = []
        for col in columns:
            val = row[col]
            if col == 'is_active' and val is not None:
                val = bool(val)
            values.append(val)
        p_cur.execute(
            f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})",
            values
        )

    print(f"{table}: {len(rows)}")
    p_cur.close()
    s_cur.close()

try:
    for table, columns in TABLES:
        copy_table(table, columns)

    cur = pg_conn.cursor()
    cur.execute("SELECT setval(pg_get_serial_sequence('sectors', 'id'), COALESCE(MAX(id), 1), true) FROM sectors;")
    cur.execute("SELECT setval(pg_get_serial_sequence('layouts', 'id'), COALESCE(MAX(id), 1), true) FROM layouts;")
    cur.execute("SELECT setval(pg_get_serial_sequence('racks', 'id'), COALESCE(MAX(id), 1), true) FROM racks;")
    cur.execute("SELECT setval(pg_get_serial_sequence('rack_slots', 'id'), COALESCE(MAX(id), 1), true) FROM rack_slots;")
    cur.execute("SELECT setval(pg_get_serial_sequence('items', 'id'), COALESCE(MAX(id), 1), true) FROM items;")
    cur.execute("SELECT setval(pg_get_serial_sequence('item_history', 'id'), COALESCE(MAX(id), 1), true) FROM item_history;")
    cur.close()

    pg_conn.commit()
    print("OK: migracja danych zakończona.")
except Exception as e:
    pg_conn.rollback()
    print("BŁĄD:", e)
    raise
finally:
    sqlite_conn.close()
    pg_conn.close()

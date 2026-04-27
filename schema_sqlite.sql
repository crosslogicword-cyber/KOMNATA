CREATE TABLE sectors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
CREATE TABLE sqlite_sequence(name,seq);
CREATE TABLE racks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            , layout_id INTEGER);
CREATE TABLE items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_name TEXT NOT NULL,
                quantity TEXT,
                notes TEXT,
                sector_id INTEGER NOT NULL,
                rack_id INTEGER NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, slot_id INTEGER, extra_field TEXT, updated_at TEXT,
                FOREIGN KEY(sector_id) REFERENCES sectors(id),
                FOREIGN KEY(rack_id) REFERENCES racks(id)
            );
CREATE TABLE rack_slots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rack_id INTEGER NOT NULL,
                code TEXT NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, sort_order INTEGER NOT NULL DEFAULT 0,
                UNIQUE(rack_id, code),
                FOREIGN KEY(rack_id) REFERENCES racks(id)
            );
CREATE TABLE layouts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                start_number INTEGER NOT NULL,
                end_number INTEGER NOT NULL,
                letters TEXT NOT NULL,
                description TEXT,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
CREATE TABLE item_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER,
                action TEXT NOT NULL,
                details TEXT,
                changed_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
CREATE TRIGGER auto_create_rack_from_layout
AFTER INSERT ON layouts
BEGIN
  INSERT OR IGNORE INTO racks (name, layout_id, is_active)
  VALUES (NEW.name, NEW.id, 1);
END;

import sqlite3
import os
from datetime import date

DB_FILE = "warehouse.db"


def get_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    c = conn.cursor()

    c.executescript('''
    CREATE TABLE IF NOT EXISTS organizations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inn VARCHAR(10) NOT NULL,
        name VARCHAR(255) NOT NULL
    );

    CREATE TABLE IF NOT EXISTS warehouses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name VARCHAR(255) NOT NULL
    );

    CREATE TABLE IF NOT EXISTS nomenclature (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name VARCHAR(255) NOT NULL,
        unit VARCHAR(255) NOT NULL
    );

    CREATE TABLE IF NOT EXISTS purchases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date DATE NOT NULL,
        supplier_id INTEGER NOT NULL,
        nomenclature_id INTEGER NOT NULL,
        warehouse_id INTEGER NOT NULL,
        quantity NUMERIC NOT NULL,
        FOREIGN KEY (supplier_id) REFERENCES organizations(id),
        FOREIGN KEY (nomenclature_id) REFERENCES nomenclature(id),
        FOREIGN KEY (warehouse_id) REFERENCES warehouses(id)
    );

    CREATE TABLE IF NOT EXISTS movements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date DATE NOT NULL,
        from_warehouse_id INTEGER NOT NULL,
        to_warehouse_id INTEGER NOT NULL,
        nomenclature_id INTEGER NOT NULL,
        quantity NUMERIC NOT NULL,
        FOREIGN KEY (from_warehouse_id) REFERENCES warehouses(id),
        FOREIGN KEY (to_warehouse_id) REFERENCES warehouses(id),
        FOREIGN KEY (nomenclature_id) REFERENCES nomenclature(id)
    );

    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date DATE NOT NULL,
        customer_id INTEGER NOT NULL,
        warehouse_id INTEGER NOT NULL,
        nomenclature_id INTEGER NOT NULL,
        quantity NUMERIC NOT NULL,
        FOREIGN KEY (customer_id) REFERENCES organizations(id),
        FOREIGN KEY (warehouse_id) REFERENCES warehouses(id),
        FOREIGN KEY (nomenclature_id) REFERENCES nomenclature(id)
    );

    CREATE TABLE IF NOT EXISTS inventory_register (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date DATE NOT NULL,
        document_type VARCHAR(255) NOT NULL,
        document_id INTEGER NOT NULL,
        nomenclature_id INTEGER NOT NULL,
        warehouse_id INTEGER NOT NULL,
        quantity NUMERIC NOT NULL,
        FOREIGN KEY (nomenclature_id) REFERENCES nomenclature(id),
        FOREIGN KEY (warehouse_id) REFERENCES warehouses(id)
    );
    ''')

    conn.commit()
    conn.close()


# ---------- Справочники ----------

def get_all(table):
    conn = get_connection()
    rows = conn.execute(f"SELECT * FROM {table} ORDER BY name").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_record(table, **kwargs):
    conn = get_connection()
    cols = ', '.join(kwargs.keys())
    placeholders = ', '.join(['?'] * len(kwargs))
    conn.execute(f"INSERT INTO {table} ({cols}) VALUES ({placeholders})", list(kwargs.values()))
    conn.commit()
    conn.close()


def update_record(table, record_id, **kwargs):
    conn = get_connection()
    sets = ', '.join([f"{k} = ?" for k in kwargs.keys()])
    conn.execute(f"UPDATE {table} SET {sets} WHERE id = ?", list(kwargs.values()) + [record_id])
    conn.commit()
    conn.close()


def delete_record(table, record_id):
    conn = get_connection()
    conn.execute(f"DELETE FROM {table} WHERE id = ?", (record_id,))
    conn.commit()
    conn.close()


# ---------- Закупки ----------

def add_purchase(date_val, supplier_id, nomenclature_id, warehouse_id, quantity):
    conn = get_connection()
    c = conn.cursor()
    c.execute(
        "INSERT INTO purchases (date, supplier_id, nomenclature_id, warehouse_id, quantity) VALUES (?,?,?,?,?)",
        (date_val, supplier_id, nomenclature_id, warehouse_id, quantity)
    )
    doc_id = c.lastrowid
    c.execute(
        "INSERT INTO inventory_register (date, document_type, document_id, nomenclature_id, warehouse_id, quantity) VALUES (?,?,?,?,?,?)",
        (date_val, 'purchase', doc_id, nomenclature_id, warehouse_id, quantity)
    )
    conn.commit()
    conn.close()


def get_purchases():
    conn = get_connection()
    rows = conn.execute('''
        SELECT p.id, p.date, o.name AS supplier, n.name AS nomenclature, n.unit,
               w.name AS warehouse, p.quantity
        FROM purchases p
        JOIN organizations o ON o.id = p.supplier_id
        JOIN nomenclature n ON n.id = p.nomenclature_id
        JOIN warehouses w ON w.id = p.warehouse_id
        ORDER BY p.date DESC, p.id DESC
    ''').fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_purchase(record_id, date_val, supplier_id, nomenclature_id, warehouse_id, quantity):
    conn = get_connection()
    conn.execute(
        "UPDATE purchases SET date=?, supplier_id=?, nomenclature_id=?, warehouse_id=?, quantity=? WHERE id=?",
        (date_val, supplier_id, nomenclature_id, warehouse_id, quantity, record_id)
    )
    conn.execute(
        "UPDATE inventory_register SET date=?, nomenclature_id=?, warehouse_id=?, quantity=? WHERE document_type='purchase' AND document_id=?",
        (date_val, nomenclature_id, warehouse_id, quantity, record_id)
    )
    conn.commit()
    conn.close()


def delete_purchase(record_id):
    conn = get_connection()
    conn.execute("DELETE FROM inventory_register WHERE document_type='purchase' AND document_id=?", (record_id,))
    conn.execute("DELETE FROM purchases WHERE id=?", (record_id,))
    conn.commit()
    conn.close()


# ---------- Продажи ----------

def add_sale(date_val, customer_id, nomenclature_id, warehouse_id, quantity):
    conn = get_connection()
    c = conn.cursor()
    c.execute(
        "INSERT INTO sales (date, customer_id, nomenclature_id, warehouse_id, quantity) VALUES (?,?,?,?,?)",
        (date_val, customer_id, nomenclature_id, warehouse_id, quantity)
    )
    doc_id = c.lastrowid
    c.execute(
        "INSERT INTO inventory_register (date, document_type, document_id, nomenclature_id, warehouse_id, quantity) VALUES (?,?,?,?,?,?)",
        (date_val, 'sale', doc_id, nomenclature_id, warehouse_id, -quantity)
    )
    conn.commit()
    conn.close()


def get_sales():
    conn = get_connection()
    rows = conn.execute('''
        SELECT s.id, s.date, o.name AS customer, n.name AS nomenclature, n.unit,
               w.name AS warehouse, s.quantity
        FROM sales s
        JOIN organizations o ON o.id = s.customer_id
        JOIN nomenclature n ON n.id = s.nomenclature_id
        JOIN warehouses w ON w.id = s.warehouse_id
        ORDER BY s.date DESC, s.id DESC
    ''').fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_sale(record_id, date_val, customer_id, nomenclature_id, warehouse_id, quantity):
    conn = get_connection()
    conn.execute(
        "UPDATE sales SET date=?, customer_id=?, nomenclature_id=?, warehouse_id=?, quantity=? WHERE id=?",
        (date_val, customer_id, nomenclature_id, warehouse_id, quantity, record_id)
    )
    conn.execute(
        "UPDATE inventory_register SET date=?, nomenclature_id=?, warehouse_id=?, quantity=? WHERE document_type='sale' AND document_id=?",
        (date_val, nomenclature_id, warehouse_id, -quantity, record_id)
    )
    conn.commit()
    conn.close()


def delete_sale(record_id):
    conn = get_connection()
    conn.execute("DELETE FROM inventory_register WHERE document_type='sale' AND document_id=?", (record_id,))
    conn.execute("DELETE FROM sales WHERE id=?", (record_id,))
    conn.commit()
    conn.close()


# ---------- Перемещения ----------

def add_movement(date_val, from_warehouse_id, to_warehouse_id, nomenclature_id, quantity):
    conn = get_connection()
    c = conn.cursor()
    c.execute(
        "INSERT INTO movements (date, from_warehouse_id, to_warehouse_id, nomenclature_id, quantity) VALUES (?,?,?,?,?)",
        (date_val, from_warehouse_id, to_warehouse_id, nomenclature_id, quantity)
    )
    doc_id = c.lastrowid
    # Списание с исходного склада
    c.execute(
        "INSERT INTO inventory_register (date, document_type, document_id, nomenclature_id, warehouse_id, quantity) VALUES (?,?,?,?,?,?)",
        (date_val, 'movement_out', doc_id, nomenclature_id, from_warehouse_id, -quantity)
    )
    # Поступление на целевой склад
    c.execute(
        "INSERT INTO inventory_register (date, document_type, document_id, nomenclature_id, warehouse_id, quantity) VALUES (?,?,?,?,?,?)",
        (date_val, 'movement_in', doc_id, nomenclature_id, to_warehouse_id, quantity)
    )
    conn.commit()
    conn.close()


def get_movements():
    conn = get_connection()
    rows = conn.execute('''
        SELECT m.id, m.date, w1.name AS from_warehouse, w2.name AS to_warehouse,
               n.name AS nomenclature, n.unit, m.quantity
        FROM movements m
        JOIN warehouses w1 ON w1.id = m.from_warehouse_id
        JOIN warehouses w2 ON w2.id = m.to_warehouse_id
        JOIN nomenclature n ON n.id = m.nomenclature_id
        ORDER BY m.date DESC, m.id DESC
    ''').fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_movement(record_id, date_val, from_warehouse_id, to_warehouse_id, nomenclature_id, quantity):
    conn = get_connection()
    conn.execute(
        "UPDATE movements SET date=?, from_warehouse_id=?, to_warehouse_id=?, nomenclature_id=?, quantity=? WHERE id=?",
        (date_val, from_warehouse_id, to_warehouse_id, nomenclature_id, quantity, record_id)
    )
    conn.execute(
        "UPDATE inventory_register SET date=?, nomenclature_id=?, warehouse_id=?, quantity=? WHERE document_type='movement_out' AND document_id=?",
        (date_val, nomenclature_id, from_warehouse_id, -quantity, record_id)
    )
    conn.execute(
        "UPDATE inventory_register SET date=?, nomenclature_id=?, warehouse_id=?, quantity=? WHERE document_type='movement_in' AND document_id=?",
        (date_val, nomenclature_id, to_warehouse_id, quantity, record_id)
    )
    conn.commit()
    conn.close()


def delete_movement(record_id):
    conn = get_connection()
    conn.execute("DELETE FROM inventory_register WHERE document_type IN ('movement_out','movement_in') AND document_id=?", (record_id,))
    conn.execute("DELETE FROM movements WHERE id=?", (record_id,))
    conn.commit()
    conn.close()


# ---------- Отчёты ----------

def get_stock_report(warehouse_id=None, nomenclature_id=None):
    conn = get_connection()
    query = '''
        SELECT n.name AS nomenclature, n.unit, w.name AS warehouse,
               SUM(r.quantity) AS balance
        FROM inventory_register r
        JOIN nomenclature n ON n.id = r.nomenclature_id
        JOIN warehouses w ON w.id = r.warehouse_id
        WHERE 1=1
    '''
    params = []
    if warehouse_id:
        query += " AND r.warehouse_id = ?"
        params.append(warehouse_id)
    if nomenclature_id:
        query += " AND r.nomenclature_id = ?"
        params.append(nomenclature_id)
    query += " GROUP BY r.nomenclature_id, r.warehouse_id HAVING SUM(r.quantity) != 0 ORDER BY n.name, w.name"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_movement_history(date_from=None, date_to=None, warehouse_id=None, nomenclature_id=None):
    conn = get_connection()
    query = '''
        SELECT r.date, r.document_type, n.name AS nomenclature, n.unit,
               w.name AS warehouse, r.quantity
        FROM inventory_register r
        JOIN nomenclature n ON n.id = r.nomenclature_id
        JOIN warehouses w ON w.id = r.warehouse_id
        WHERE 1=1
    '''
    params = []
    if date_from:
        query += " AND r.date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND r.date <= ?"
        params.append(date_to)
    if warehouse_id:
        query += " AND r.warehouse_id = ?"
        params.append(warehouse_id)
    if nomenclature_id:
        query += " AND r.nomenclature_id = ?"
        params.append(nomenclature_id)
    query += " ORDER BY r.date DESC, r.id DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]

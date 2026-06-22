"""
Migración v2: agrega columnas a asignaturas y crea tablas bibliografia/linkografia.
Idempotente: usa IF NOT EXISTS / ignora errores de columna duplicada.
"""
import sqlite3
from pathlib import Path

RUTA_DB = Path("data/sistema.db")

NUEVAS_COLUMNAS = [
    "ALTER TABLE asignaturas ADD COLUMN horas_directa REAL",
    "ALTER TABLE asignaturas ADD COLUMN horas_autonoma REAL",
    "ALTER TABLE asignaturas ADD COLUMN semanas INTEGER",
    "ALTER TABLE asignaturas ADD COLUMN creditos INTEGER",
    "ALTER TABLE asignaturas ADD COLUMN descripcion TEXT",
    "ALTER TABLE asignaturas ADD COLUMN otros_recursos TEXT",
]

NUEVAS_TABLAS = """
CREATE TABLE IF NOT EXISTS bibliografia (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id INTEGER NOT NULL,
    tipo TEXT NOT NULL,
    numero TEXT, autor TEXT, titulo TEXT, editorial TEXT,
    anio TEXT, isbn TEXT, ejemplares TEXT,
    FOREIGN KEY (asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS linkografia (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id INTEGER NOT NULL,
    tipo_doc TEXT, autor TEXT, titulo_articulo TEXT, anio TEXT,
    titulo_revista TEXT, volumen TEXT, url TEXT, disponible TEXT,
    FOREIGN KEY (asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_biblio_asig ON bibliografia(asignatura_id);
CREATE INDEX IF NOT EXISTS idx_link_asig ON linkografia(asignatura_id);
"""

def migrar():
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")

    print("Migrando schema v2...")

    for sql in NUEVAS_COLUMNAS:
        try:
            conn.execute(sql)
            print(f"  OK: {sql}")
        except sqlite3.OperationalError as e:
            if "duplicate column name" in str(e):
                print(f"  YA EXISTE: {sql.split()[-2]}")
            else:
                print(f"  ERROR: {e}")

    conn.executescript(NUEVAS_TABLAS)
    print("  OK: tablas bibliografia, linkografia e índices creados (IF NOT EXISTS)")

    conn.commit()
    conn.close()

    print("\nMigración completada.")

    # Verificar
    conn2 = sqlite3.connect(str(RUTA_DB))
    cols = [r[1] for r in conn2.execute("PRAGMA table_info(asignaturas)").fetchall()]
    print(f"\nColumnas de asignaturas: {cols}")
    tables = [r[0] for r in conn2.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
    print(f"Tablas: {tables}")
    conn2.close()

if __name__ == "__main__":
    migrar()

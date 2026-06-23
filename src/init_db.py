"""
init_db.py — Inicializa sistema.db desde cero con esquema normalizado.

Lee data/niveles_dominio.csv y data/resultados_aprendizaje.csv para
poblar las tablas maestras. El resto de tablas quedan vacías listas
para cargarse con cargar_todos.py.

Uso:
    python3 src/init_db.py            # crea/recrea data/sistema.db
    python3 src/init_db.py --verify   # solo muestra estado de la BD
"""

import csv
import re
import sqlite3
import sys
from pathlib import Path

RUTA_DB              = Path("data/sistema.db")
CSV_NIVELES          = Path("data/niveles_dominio.csv")
CSV_RAS              = Path("data/resultados_aprendizaje.csv")

# Tipo de cada competencia según su prefijo
_TIPO = {
    "CL": "licenciatura",
    "CE": "titulo",
    "CG": "sello_uv",
}


# ── DDL ───────────────────────────────────────────────────────────────────────

SCHEMA = """
-- ============================================================
--  TABLAS MAESTRAS
-- ============================================================

CREATE TABLE IF NOT EXISTS competencias (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo      TEXT    NOT NULL UNIQUE,
    tipo        TEXT    NOT NULL CHECK(tipo IN ('licenciatura','titulo','sello_uv')),
    descripcion TEXT    DEFAULT ''
);

CREATE TABLE IF NOT EXISTS niveles_dominio (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    competencia_id  INTEGER NOT NULL REFERENCES competencias(id) ON DELETE CASCADE,
    codigo_nivel    TEXT    NOT NULL,   -- ND1, ND2, ND3
    descripcion     TEXT    DEFAULT '',
    UNIQUE (competencia_id, codigo_nivel)
);

CREATE TABLE IF NOT EXISTS resultados_aprendizaje (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    nivel_dominio_id INTEGER NOT NULL REFERENCES niveles_dominio(id) ON DELETE CASCADE,
    codigo_ra       TEXT    NOT NULL,   -- RA1, RA2 ... / D1, D2 ...
    codigo_completo TEXT    NOT NULL UNIQUE,  -- ej. "CE1, ND2, RA3"
    descripcion     TEXT    DEFAULT ''
);

-- ============================================================
--  ASIGNATURAS Y DATOS DEL PROGRAMA
-- ============================================================

CREATE TABLE IF NOT EXISTS asignaturas (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo          TEXT    NOT NULL UNIQUE,
    nombre          TEXT    NOT NULL DEFAULT '',
    semestre        INTEGER,
    nivel           TEXT    DEFAULT '',
    duracion        TEXT    DEFAULT '',
    tipo            TEXT    DEFAULT 'disciplinar',
    facultad        TEXT    DEFAULT '',
    carrera         TEXT    DEFAULT '',
    requisitos      TEXT    DEFAULT '',
    horas_directa   REAL,
    horas_autonoma  REAL,
    semanas         INTEGER,
    creditos        INTEGER,
    descripcion     TEXT    DEFAULT '',
    otros_recursos  TEXT    DEFAULT '',
    version         TEXT    DEFAULT '',
    archivo_origen  TEXT    DEFAULT ''
);

CREATE TABLE IF NOT EXISTS responsables (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id   INTEGER NOT NULL REFERENCES asignaturas(id) ON DELETE CASCADE,
    rol             TEXT    NOT NULL,   -- 'responsable' | 'docente_a_cargo'
    nombre          TEXT    NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS unidades (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id   INTEGER NOT NULL REFERENCES asignaturas(id) ON DELETE CASCADE,
    orden           INTEGER,
    nombre          TEXT    DEFAULT '',
    contenidos      TEXT    DEFAULT '',
    indicador_logro TEXT    DEFAULT ''
);

CREATE TABLE IF NOT EXISTS metodologias (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id   INTEGER NOT NULL REFERENCES asignaturas(id) ON DELETE CASCADE,
    descripcion     TEXT    NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS evaluaciones (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id   INTEGER NOT NULL REFERENCES asignaturas(id) ON DELETE CASCADE,
    tipo            TEXT    NOT NULL DEFAULT '',
    porcentaje      TEXT    DEFAULT ''
);

CREATE TABLE IF NOT EXISTS bibliografia (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id   INTEGER NOT NULL REFERENCES asignaturas(id) ON DELETE CASCADE,
    tipo            TEXT    NOT NULL CHECK(tipo IN ('basica','complementaria')),
    numero          TEXT    DEFAULT '',
    autor           TEXT    DEFAULT '',
    titulo          TEXT    DEFAULT '',
    editorial       TEXT    DEFAULT '',
    anio            TEXT    DEFAULT '',
    isbn            TEXT    DEFAULT '',
    ejemplares      TEXT    DEFAULT '',
    referencia      TEXT    DEFAULT ''
);

CREATE TABLE IF NOT EXISTS linkografia (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id   INTEGER NOT NULL REFERENCES asignaturas(id) ON DELETE CASCADE,
    descripcion     TEXT    DEFAULT '',
    url             TEXT    DEFAULT '',
    titulo_articulo TEXT    DEFAULT ''
);

-- ============================================================
--  TRIBUTACIONES (N:M  asignaturas ↔ resultados_aprendizaje)
-- ============================================================

CREATE TABLE IF NOT EXISTS tributaciones (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id   INTEGER NOT NULL REFERENCES asignaturas(id)          ON DELETE CASCADE,
    ra_id           INTEGER NOT NULL REFERENCES resultados_aprendizaje(id) ON DELETE CASCADE,
    UNIQUE (asignatura_id, ra_id)
);

-- ============================================================
--  ÍNDICES
-- ============================================================

CREATE INDEX IF NOT EXISTS idx_nd_comp     ON niveles_dominio(competencia_id);
CREATE INDEX IF NOT EXISTS idx_ra_nd       ON resultados_aprendizaje(nivel_dominio_id);
CREATE INDEX IF NOT EXISTS idx_trib_asig   ON tributaciones(asignatura_id);
CREATE INDEX IF NOT EXISTS idx_trib_ra     ON tributaciones(ra_id);
CREATE INDEX IF NOT EXISTS idx_unid_asig   ON unidades(asignatura_id);
CREATE INDEX IF NOT EXISTS idx_biblio_asig ON bibliografia(asignatura_id);
CREATE INDEX IF NOT EXISTS idx_link_asig   ON linkografia(asignatura_id);
"""


# ── Helpers ───────────────────────────────────────────────────────────────────

def _tipo_comp(codigo: str) -> str:
    prefijo = re.match(r"[A-Z]+", codigo).group()
    return _TIPO.get(prefijo, "titulo")


def _leer_csv(ruta: Path) -> list[dict]:
    with open(ruta, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f, delimiter=";"))


# ── Población de tablas maestras ─────────────────────────────────────────────

def poblar_maestras(conn: sqlite3.Connection) -> None:
    niveles_rows = _leer_csv(CSV_NIVELES)
    ras_rows     = _leer_csv(CSV_RAS)

    # 1. Competencias (únicas, inferidas de ambos CSV)
    codigos_comp = {r["competencia"].strip() for r in niveles_rows} | \
                   {r["competencia"].strip() for r in ras_rows}

    for cod in sorted(codigos_comp):
        conn.execute(
            "INSERT OR IGNORE INTO competencias (codigo, tipo) VALUES (?, ?)",
            (cod, _tipo_comp(cod))
        )

    def _comp_id(cod):
        return conn.execute(
            "SELECT id FROM competencias WHERE codigo = ?", (cod,)
        ).fetchone()[0]

    # 2. Niveles de Dominio
    for r in niveles_rows:
        cod_comp = r["competencia"].strip()
        conn.execute(
            """INSERT OR REPLACE INTO niveles_dominio
               (competencia_id, codigo_nivel, descripcion)
               VALUES (?, ?, ?)""",
            (_comp_id(cod_comp), r["codigo_nivel"].strip(), r["descripcion"].strip())
        )

    def _nd_id(cod_comp, cod_nivel):
        return conn.execute(
            """SELECT nd.id FROM niveles_dominio nd
               JOIN competencias c ON c.id = nd.competencia_id
               WHERE c.codigo = ? AND nd.codigo_nivel = ?""",
            (cod_comp, cod_nivel)
        ).fetchone()[0]

    # 3. Resultados de Aprendizaje
    for r in ras_rows:
        cod_comp  = r["competencia"].strip()
        cod_nivel = r["nivel_dominio"].strip()
        cod_ra    = r["codigo_ra"].strip()
        desc      = r["descripcion"].strip()
        codigo_completo = f"{cod_comp}, {cod_nivel}, {cod_ra}"

        conn.execute(
            """INSERT OR REPLACE INTO resultados_aprendizaje
               (nivel_dominio_id, codigo_ra, codigo_completo, descripcion)
               VALUES (?, ?, ?, ?)""",
            (_nd_id(cod_comp, cod_nivel), cod_ra, codigo_completo, desc)
        )


# ── Verificación ──────────────────────────────────────────────────────────────

def verificar(conn: sqlite3.Connection) -> None:
    tablas = [r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
    ).fetchall()]
    print("\n=== Tablas en sistema.db ===")
    for t in tablas:
        n = conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        print(f"  {t:<30} {n:>5} filas")

    print("\n=== Competencias ===")
    for r in conn.execute("SELECT codigo, tipo FROM competencias ORDER BY tipo, codigo"):
        print(f"  {r[0]:6}  {r[1]}")

    print("\n=== Niveles de Dominio por competencia ===")
    for r in conn.execute("""
        SELECT c.codigo, nd.codigo_nivel, substr(nd.descripcion,1,60)
        FROM niveles_dominio nd JOIN competencias c ON c.id=nd.competencia_id
        ORDER BY c.codigo, nd.codigo_nivel
    """):
        print(f"  {r[0]:6} {r[1]:4}  {r[2]}…")

    n_ra = conn.execute("SELECT COUNT(*) FROM resultados_aprendizaje").fetchone()[0]
    print(f"\n  Total RAs: {n_ra}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    solo_verificar = "--verify" in sys.argv

    if solo_verificar:
        if not RUTA_DB.exists():
            print(f"ERROR: {RUTA_DB} no existe. Ejecuta sin --verify primero.")
            sys.exit(1)
        conn = sqlite3.connect(str(RUTA_DB))
        conn.execute("PRAGMA foreign_keys = ON")
        verificar(conn)
        conn.close()
        return

    # Recrear BD
    if RUTA_DB.exists():
        RUTA_DB.unlink()
        print(f"BD anterior eliminada: {RUTA_DB}")

    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA)
    print("Esquema creado.")

    poblar_maestras(conn)
    conn.commit()
    print("Tablas maestras pobladas desde CSV.")

    verificar(conn)
    conn.close()
    print(f"\nBD lista en: {RUTA_DB}")


if __name__ == "__main__":
    main()

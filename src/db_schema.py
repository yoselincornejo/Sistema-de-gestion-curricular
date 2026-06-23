"""
Define el schema de la BD y la inicializa.
Si la BD existe, no hace nada destructivo (CREATE IF NOT EXISTS).
Para recargar desde cero, borrar data/sistema.db manualmente.
"""

import sqlite3
from pathlib import Path

RUTA_DB = Path("data/sistema.db")

SCHEMA_SQL = """
-- Programas académicos / asignaturas
CREATE TABLE IF NOT EXISTS asignaturas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT NOT NULL UNIQUE,
    nombre TEXT NOT NULL,
    semestre INTEGER,
    nivel TEXT,
    duracion TEXT,
    tipo TEXT,                  -- disciplinar / transversal / flexible / integradora / anexo
    facultad TEXT,
    carrera TEXT,
    requisitos TEXT,
    version TEXT,
    archivo_origen TEXT
);

-- Responsables (puede haber varios por asignatura: responsable y docente_a_cargo)
CREATE TABLE IF NOT EXISTS responsables (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id INTEGER NOT NULL,
    rol TEXT NOT NULL,          -- 'responsable' o 'docente_a_cargo'
    nombre TEXT,
    FOREIGN KEY (asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
);

-- Competencias del perfil de egreso (8 totales en plan 2025)
CREATE TABLE IF NOT EXISTS competencias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT NOT NULL UNIQUE,        -- CL1, CL2, CE1, CE2, CG1, CG2, CG3, CG4
    tipo TEXT NOT NULL,                 -- 'licenciatura' / 'titulo' / 'sello_uv'
    descripcion TEXT
);

-- Niveles de dominio por competencia (ND1, ND2, ND3)
CREATE TABLE IF NOT EXISTS niveles_dominio (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    competencia_id INTEGER NOT NULL,
    codigo_nivel TEXT NOT NULL,          -- ND1, ND2, ND3
    descripcion TEXT,
    FOREIGN KEY (competencia_id) REFERENCES competencias(id) ON DELETE CASCADE,
    UNIQUE (competencia_id, codigo_nivel)
);

-- Resultados de aprendizaje (RA) catalogados por competencia y nivel
CREATE TABLE IF NOT EXISTS resultados_aprendizaje (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    competencia_id INTEGER NOT NULL,
    nivel_dominio TEXT,                 -- ND1, ND2, ND3
    codigo TEXT NOT NULL,               -- RA1, RA2, D1, D2, etc.
    codigo_completo TEXT NOT NULL UNIQUE,   -- "CL2, ND2, RA1" para búsqueda rápida
    descripcion TEXT,
    FOREIGN KEY (competencia_id) REFERENCES competencias(id) ON DELETE CASCADE
);

-- Tributación: qué asignatura aporta a qué RA (corazón del sistema)
CREATE TABLE IF NOT EXISTS tributaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id INTEGER NOT NULL,
    ra_id INTEGER NOT NULL,
    FOREIGN KEY (asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE,
    FOREIGN KEY (ra_id) REFERENCES resultados_aprendizaje(id) ON DELETE CASCADE,
    UNIQUE (asignatura_id, ra_id)
);

-- Unidades temáticas de cada asignatura
CREATE TABLE IF NOT EXISTS unidades (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id INTEGER NOT NULL,
    orden INTEGER,                       -- número de unidad
    nombre TEXT,
    contenidos TEXT,
    indicador_logro TEXT,
    FOREIGN KEY (asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
);

-- Metodologías (puede haber varias por asignatura)
CREATE TABLE IF NOT EXISTS metodologias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id INTEGER NOT NULL,
    descripcion TEXT,
    FOREIGN KEY (asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
);

-- Evaluaciones (puede haber varias por asignatura)
CREATE TABLE IF NOT EXISTS evaluaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    asignatura_id INTEGER NOT NULL,
    tipo TEXT,
    porcentaje TEXT,
    FOREIGN KEY (asignatura_id) REFERENCES asignaturas(id) ON DELETE CASCADE
);

-- Índices para queries comunes
CREATE INDEX IF NOT EXISTS idx_asig_codigo ON asignaturas(codigo);
CREATE INDEX IF NOT EXISTS idx_asig_semestre ON asignaturas(semestre);
CREATE INDEX IF NOT EXISTS idx_ra_competencia ON resultados_aprendizaje(competencia_id);
CREATE INDEX IF NOT EXISTS idx_asig_ra_asignatura ON tributaciones(asignatura_id);
CREATE INDEX IF NOT EXISTS idx_asig_ra_ra ON tributaciones(ra_id);
"""

# Seed inicial de las 8 competencias del plan 2025
COMPETENCIAS_BASE = [
    ("CL1", "licenciatura", "Competencia de Licenciatura 1"),
    ("CL2", "licenciatura", "Competencia de Licenciatura 2"),
    ("CE1", "titulo", "Competencia Específica del Título 1"),
    ("CE2", "titulo", "Competencia Específica del Título 2"),
    ("CG1", "sello_uv", "Competencia Genérica Sello UV 1"),
    ("CG2", "sello_uv", "Competencia Genérica Sello UV 2"),
    ("CG3", "sello_uv", "Competencia Genérica Sello UV 3"),
    ("CG4", "sello_uv", "Competencia Genérica Sello UV 4"),
]


def crear_schema(ruta_db=RUTA_DB):
    ruta_db.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(ruta_db))
    conn.execute("PRAGMA foreign_keys = ON")
    conn.executescript(SCHEMA_SQL)

    # Insertar competencias base si no existen
    for codigo, tipo, descripcion in COMPETENCIAS_BASE:
        conn.execute(
            "INSERT OR IGNORE INTO competencias (codigo, tipo, descripcion) VALUES (?, ?, ?)",
            (codigo, tipo, descripcion)
        )

    conn.commit()
    conn.close()

    print(f"Schema creado en: {ruta_db}")
    print(f"Competencias base insertadas: {len(COMPETENCIAS_BASE)}")


if __name__ == "__main__":
    crear_schema()
"""
cargar_todos.py — Carga masiva de programas .docx a la BD.
Traversa data/programas/ recursivamente, parsea y carga.
"""
import sys
import sqlite3
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))

from parsear_programas import parsear_docx
from cargar_a_db import cargar_programa, RUTA_DB

CARPETA_PROGRAMAS = Path("data/programas")

# Asignatura a NO tocar
EXCLUIDOS = {"IMAT 423"}


def cargar_todos():
    # Verificar y ejecutar migración si es necesario
    conn = sqlite3.connect(str(RUTA_DB))
    conn.execute("PRAGMA foreign_keys = ON")

    # Verificar si las nuevas tablas existen
    tablas = [r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    ).fetchall()]

    if "bibliografia" not in tablas:
        print("Ejecutando migración de schema v2...")
        # Add columns one by one to handle errors
        for sql in [
            "ALTER TABLE asignaturas ADD COLUMN horas_directa REAL",
            "ALTER TABLE asignaturas ADD COLUMN horas_autonoma REAL",
            "ALTER TABLE asignaturas ADD COLUMN semanas INTEGER",
            "ALTER TABLE asignaturas ADD COLUMN creditos INTEGER",
            "ALTER TABLE asignaturas ADD COLUMN descripcion TEXT",
            "ALTER TABLE asignaturas ADD COLUMN otros_recursos TEXT",
        ]:
            try:
                conn.execute(sql)
            except sqlite3.OperationalError as e:
                if "duplicate column name" not in str(e):
                    print(f"  ERROR: {e}")
        conn.executescript("""
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
        """)
        conn.commit()
        print("Migración completada.")

    # Encontrar todos los .docx
    docx_files = sorted(CARPETA_PROGRAMAS.rglob("*.docx"))
    print(f"\nEncontrados {len(docx_files)} archivos .docx en {CARPETA_PROGRAMAS}\n")

    procesados = 0
    errores = []
    ras_no_resueltos_total = []
    excluidos = 0

    for docx_path in docx_files:
        nombre_archivo = docx_path.name

        try:
            # Parsear
            programa = parsear_docx(str(docx_path))
            codigo = programa.get("identificacion", {}).get("codigo", "").strip()

            # Verificar exclusiones
            if any(exc in codigo for exc in EXCLUIDOS):
                print(f"  SKIP {codigo:15} {nombre_archivo[:50]} (excluido)")
                excluidos += 1
                continue

            # Cargar
            result = cargar_programa(conn, programa)
            if len(result) == 3:
                asig_id, estado, ras_nr = result
            else:
                asig_id, estado = result
                ras_nr = []

            if estado == "ok":
                procesados += 1
                nombre = programa.get("identificacion", {}).get("nombre", "")[:45]
                semestre = programa.get("semestre", "?")
                creditos = programa.get("identificacion", {}).get("creditos", "?")
                n_unidades = len(programa.get("unidades", []))
                n_biblio = len(programa.get("bibliografia_basica", [])) + len(programa.get("bibliografia_complementaria", []))
                print(f"  OK  {codigo:15} S{semestre} {nombre:45} | U:{n_unidades} B:{n_biblio} C:{creditos}")
                if ras_nr:
                    print(f"      RAs no resueltos: {ras_nr}")
                    ras_no_resueltos_total.extend([(codigo, r) for r in ras_nr])
            else:
                errores.append((nombre_archivo, estado))
                print(f"  --  {nombre_archivo[:55]} ({estado})")

        except Exception as e:
            errores.append((nombre_archivo, str(e)))
            print(f"  XX  {nombre_archivo[:55]} ERROR: {e}")

    conn.commit()
    conn.close()

    print(f"\n{'='*70}")
    print(f"Documentos procesados: {procesados}/{len(docx_files)} (excluidos: {excluidos})")

    if errores:
        print(f"\nErrores ({len(errores)}):")
        for nombre, motivo in errores:
            print(f"  - {nombre}: {motivo}")

    if ras_no_resueltos_total:
        print(f"\nRAs no resueltos ({len(ras_no_resueltos_total)}):")
        for codigo, ra in ras_no_resueltos_total:
            print(f"  - {codigo}: {ra}")

    print("\n--- VERIFICACIÓN FINAL ---")
    conn2 = sqlite3.connect(str(RUTA_DB))
    print("Asignaturas:", conn2.execute("SELECT COUNT(*) FROM asignaturas").fetchone()[0])
    print("Con créditos:", conn2.execute("SELECT COUNT(*) FROM asignaturas WHERE creditos IS NOT NULL").fetchone()[0])
    print("Con descripción:", conn2.execute("SELECT COUNT(*) FROM asignaturas WHERE descripcion IS NOT NULL AND descripcion != ''").fetchone()[0])
    print("Unidades con indicador_logro:", conn2.execute("SELECT COUNT(*) FROM unidades WHERE indicador_logro IS NOT NULL AND indicador_logro != ''").fetchone()[0])
    print("Bibliografía:", conn2.execute("SELECT COUNT(*) FROM bibliografia").fetchone()[0])
    print("Tributaciones:", conn2.execute("SELECT COUNT(*) FROM tributaciones").fetchone()[0])
    conn2.close()


if __name__ == "__main__":
    cargar_todos()
